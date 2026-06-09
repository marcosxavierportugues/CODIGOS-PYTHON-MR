from pathlib import Path
from urllib.parse import urlparse, unquote
from typing import Optional
from email.message import Message
import mimetypes
import requests
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox


def limpar_nome_arquivo(nome: str) -> str:
    """
    Remove caracteres inválidos para nome de arquivo no Windows.
    """
    caracteres_invalidos = '<>:"/\\|?*'

    for caractere in caracteres_invalidos:
        nome = nome.replace(caractere, "_")

    nome = nome.strip()

    if not nome:
        return "arquivo_baixado"

    return nome


def obter_nome_pela_url(url: str) -> str:
    """
    Tenta extrair o nome do arquivo pela URL.
    """
    caminho = urlparse(url).path
    nome = Path(unquote(caminho)).name

    if not nome:
        nome = "arquivo_baixado"

    return limpar_nome_arquivo(nome)


def obter_nome_por_content_disposition(content_disposition: Optional[str]) -> Optional[str]:
    """
    Tenta extrair o nome do arquivo pelo cabeçalho Content-Disposition.
    Exemplo:
    Content-Disposition: attachment; filename="arquivo.pdf"
    """
    if not content_disposition:
        return None

    mensagem = Message()
    mensagem["content-disposition"] = content_disposition

    nome = mensagem.get_param("filename", header="content-disposition")

    if not nome:
        return None

    return limpar_nome_arquivo(str(nome))


def ajustar_extensao_por_content_type(nome_arquivo: str, content_type: Optional[str]) -> str:
    """
    Ajusta a extensão se o servidor informar o tipo do arquivo.
    Exemplo: application/pdf => .pdf
    """
    if not content_type:
        return nome_arquivo

    content_type = content_type.split(";")[0].strip().lower()

    extensao_atual = Path(nome_arquivo).suffix.lower()

    if content_type == "application/pdf" and extensao_atual != ".pdf":
        return f"{Path(nome_arquivo).stem}.pdf"

    extensao_detectada = mimetypes.guess_extension(content_type)

    if extensao_detectada and not extensao_atual:
        return f"{nome_arquivo}{extensao_detectada}"

    return nome_arquivo


def escolher_pasta_destino() -> Optional[Path]:
    """
    Abre janela para escolher a pasta de destino.
    """
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    pasta = filedialog.askdirectory(
        title="Escolha a pasta onde o arquivo será baixado"
    )

    root.destroy()

    if not pasta:
        return None

    return Path(pasta)


def pedir_link() -> Optional[str]:
    """
    Abre uma janela para colar o link.
    Evita problemas com input() no terminal.
    """
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    link = simpledialog.askstring(
        title="Link do arquivo",
        prompt="Cole o link do arquivo:"
    )

    root.destroy()

    if not link:
        return None

    link = link.strip()

    if not link:
        return None

    return link


def gerar_caminho_disponivel(caminho: Path) -> Path:
    """
    Evita sobrescrever arquivo existente.
    Exemplo:
    arquivo.pdf
    arquivo (1).pdf
    arquivo (2).pdf
    """
    if not caminho.exists():
        return caminho

    pasta = caminho.parent
    nome_base = caminho.stem
    extensao = caminho.suffix

    contador = 1

    while True:
        novo_caminho = pasta / f"{nome_base} ({contador}){extensao}"

        if not novo_caminho.exists():
            return novo_caminho

        contador += 1


def remover_temporario(caminho_temporario: Path) -> None:
    """
    Remove arquivo temporário incompleto, se existir.
    """
    try:
        if caminho_temporario.exists():
            caminho_temporario.unlink()
    except OSError:
        pass


def baixar_arquivo(
    url: str,
    pasta_destino: Path,
    nome_arquivo: Optional[str] = None,
) -> Path:
    """
    Baixa um arquivo de uma URL e salva na pasta escolhida.
    """

    if not url.startswith(("http://", "https://")):
        raise ValueError("URL inválida. Use um link começando com http:// ou https://.")

    pasta_destino.mkdir(parents=True, exist_ok=True)

    caminho_temporario: Optional[Path] = None

    try:
        with requests.get(url, stream=True, timeout=30, allow_redirects=True) as resposta:
            resposta.raise_for_status()

            if nome_arquivo is None:
                nome_cabecalho = obter_nome_por_content_disposition(
                    resposta.headers.get("Content-Disposition")
                )

                if nome_cabecalho:
                    nome_final = nome_cabecalho
                else:
                    nome_final = obter_nome_pela_url(resposta.url)
            else:
                nome_final = limpar_nome_arquivo(nome_arquivo)

            nome_final = ajustar_extensao_por_content_type(
                nome_final,
                resposta.headers.get("Content-Type")
            )

            caminho_final = pasta_destino / nome_final
            caminho_final = gerar_caminho_disponivel(caminho_final)

            caminho_temporario = caminho_final.with_name(
                caminho_final.name + ".part"
            )

            with open(caminho_temporario, "wb") as arquivo:
                for bloco in resposta.iter_content(chunk_size=1024 * 1024):
                    if bloco:
                        arquivo.write(bloco)

            caminho_temporario.rename(caminho_final)

            return caminho_final

    except requests.exceptions.Timeout as erro:
        if caminho_temporario is not None:
            remover_temporario(caminho_temporario)

        raise RuntimeError("Tempo limite excedido durante o download.") from erro

    except requests.exceptions.ConnectionError as erro:
        if caminho_temporario is not None:
            remover_temporario(caminho_temporario)

        raise RuntimeError("Erro de conexão ao tentar baixar o arquivo.") from erro

    except requests.exceptions.HTTPError as erro:
        if caminho_temporario is not None:
            remover_temporario(caminho_temporario)

        resposta_erro = erro.response

        if resposta_erro is not None:
            status_code = resposta_erro.status_code
        else:
            status_code = "desconhecido"

        raise RuntimeError(
            f"Erro HTTP ao baixar o arquivo. Código: {status_code}"
        ) from erro

    except requests.exceptions.RequestException as erro:
        if caminho_temporario is not None:
            remover_temporario(caminho_temporario)

        raise RuntimeError(f"Erro inesperado no download: {erro}") from erro

    except OSError as erro:
        if caminho_temporario is not None:
            remover_temporario(caminho_temporario)

        raise RuntimeError(f"Erro ao salvar o arquivo no disco: {erro}") from erro


def main() -> None:
    link = pedir_link()

    if link is None:
        print("Operação cancelada. Nenhum link foi informado.")
        return

    pasta = escolher_pasta_destino()

    if pasta is None:
        print("Operação cancelada. Nenhuma pasta foi escolhida.")
        return

    try:
        arquivo_salvo = baixar_arquivo(
            url=link,
            pasta_destino=pasta,
        )

        mensagem = f"Download concluído com sucesso.\n\nArquivo salvo em:\n{arquivo_salvo}"
        print(mensagem)

        messagebox.showinfo(
            title="Download concluído",
            message=mensagem
        )

    except Exception as erro:
        mensagem = f"Falha ao baixar arquivo:\n\n{erro}"
        print(mensagem)

        messagebox.showerror(
            title="Erro no download",
            message=mensagem
        )


if __name__ == "__main__":
    main()