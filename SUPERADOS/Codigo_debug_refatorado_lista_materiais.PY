# -*- coding: utf-8 -*-
"""
Lista de Materiais IFMT - Gerador Universal
Suporta: qualquer peça (pilares, vigas, lajes, sapatas...)
         qualquer numero de segmentos e pavimentos
         modo RESUMO ou COMPLETO
         nome de arquivo automatico por metadados
"""
import os
import re
import shutil
import traceback
import unicodedata
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from tkinter import Tk, filedialog, simpledialog, messagebox
import tkinter as tk

# ─── DIALOGS ─────────────────────────────────────────────────────────────────

def _root():
    """Cria uma raiz Tk oculta e sempre no topo."""
    root = Tk()
    root.withdraw()
    root.lift()
    try:
        root.attributes("-topmost", True)
    except tk.TclError:
        pass
    return root

def _confirm_cancel(parent=None):
    """Confirma cancelamento sem lançar exceção dentro de callback Tk."""
    owns_root = parent is None
    root = parent or _root()
    try:
        return bool(
            messagebox.askyesno(
                "Cancelar operação",
                "Cancelar a operação?",
                parent=root,
                icon="warning",
            )
        )
    finally:
        if owns_root:
            try:
                root.destroy()
            except tk.TclError:
                pass


def _ask_str(title, prompt, default=""):
    """Solicita texto e permite voltar ao diálogo sem recursão."""
    while True:
        root = _root()
        try:
            value = simpledialog.askstring(
                title,
                prompt,
                parent=root,
                initialvalue=default,
            )
            if value is not None:
                return str(value).strip().upper()
            if _confirm_cancel(root):
                raise SystemExit("Cancelado pelo usuário")
        finally:
            try:
                root.destroy()
            except tk.TclError:
                pass


def _ask_file(title):
    root = _root()
    try:
        selected = filedialog.askopenfilename(
            title=title,
            filetypes=[("Excel", "*.xlsx *.xls")],
            parent=root,
        )
    finally:
        try:
            root.destroy()
        except tk.TclError:
            pass
    if not selected:
        raise SystemExit(f"Cancelado: {title}")
    return Path(selected)

def _ask_save(title, initial_name):
    root = _root()
    try:
        selected = filedialog.asksaveasfilename(
            title=title,
            initialfile=initial_name,
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            parent=root,
        )
    finally:
        try:
            root.destroy()
        except tk.TclError:
            pass
    if not selected:
        raise SystemExit("Cancelado: salvar arquivo")
    return Path(selected)

def _ask_files(title):
    root = _root()
    try:
        selected = filedialog.askopenfilenames(
            title=title,
            filetypes=[("Excel", "*.xlsx *.xls")],
            parent=root,
        )
    finally:
        try:
            root.destroy()
        except tk.TclError:
            pass
    if not selected:
        raise SystemExit(f"Cancelado: {title}")
    return [Path(item) for item in selected]


def _ask_main_action():
    root = _root()
    result = [3]
    try:
        win = tk.Toplevel(root)
        win.title("O que deseja fazer?")
        win.attributes("-topmost", True)
        tk.Label(
            win,
            text="O que deseja fazer?",
            font=("Arial", 12, "bold"),
            padx=20,
            pady=12,
        ).pack()

        action = tk.IntVar(value=1)
        options = [
            ("Gerar lista de material", 1),
            ("Unir listas de materiais", 2),
            ("Cancelar", 3),
        ]
        for text, value in options:
            tk.Radiobutton(
                win,
                text=text,
                variable=action,
                value=value,
                font=("Arial", 11),
                anchor="w",
            ).pack(fill="x", padx=20, pady=2)

        def finish(value):
            result[0] = value
            win.destroy()

        frame = tk.Frame(win)
        frame.pack(pady=10)
        tk.Button(
            frame,
            text="OK",
            command=lambda: finish(action.get()),
            width=12,
        ).pack(side="left", padx=8)
        tk.Button(
            frame,
            text="Cancelar",
            command=lambda: finish(3),
            width=12,
            fg="white",
            bg="red",
        ).pack(side="left", padx=8)
        win.protocol("WM_DELETE_WINDOW", lambda: finish(3))
        win.wait_window()
        return result[0]
    finally:
        try:
            root.destroy()
        except tk.TclError:
            pass


def _normalize_text(value):
    """Normaliza texto para comparações sem alterar o valor original."""
    if value is None:
        return ""
    try:
        if bool(pd.isna(value)):
            return ""
    except (TypeError, ValueError):
        pass

    text = str(value).strip()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(
        character
        for character in text
        if not unicodedata.combining(character)
    )
    return re.sub(r"\s+", " ", text).casefold()

def _extract_element_category(value):
    """Extrai um prefixo textual apenas para avisos de divergência."""
    text = _normalize_text(value)
    text = re.sub(r"\d+", "", text)
    text = re.sub(r"[^a-z ]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text.split()[0] if text else ""

def _extract_segment_label(filename, index):
    name = str(filename)
    m = re.search(r"(SEG(?:MENTO)?\s*\d+)", name, re.IGNORECASE)
    if m:
        label = m.group(1).upper().replace("SEGMENTO", "SEG")
        return re.sub(r"\s+", " ", label).strip()
    return f"SEG {index}"


def _find_column(columns, keywords, exclude=None):
    """Localiza uma coluna por termos normalizados, respeitando exclusões."""
    normalized_keywords = [_normalize_text(item) for item in keywords]
    normalized_exclude = [_normalize_text(item) for item in (exclude or [])]

    for column in columns:
        normalized = _normalize_text(column)
        if any(term and term in normalized for term in normalized_exclude):
            continue
        if any(term and term in normalized for term in normalized_keywords):
            return column
    return None





def _column_key(value):
    """Normaliza um cabeçalho para comparação segura."""
    text = _normalize_text(value)
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def _is_altoqi_element_column(column):
    """
    Identifica especificamente a coluna técnica AltoQi/Eberick ``Elemento``.

    Exemplo aceito:
    ``AltoQi_Eberick_Elemento\nElemento``.
    """
    key = _column_key(column)
    if key in {
        "altoqi eberick elemento",
        "altoqi eberick elemento elemento",
    }:
        return True
    return (
        key.startswith("altoqi eberick elemento ")
        and key.endswith(" elemento")
    )


def _drop_ignored_source_columns(df):
    """Exclui colunas técnicas ignoradas e todos os seus valores."""
    ignored = [c for c in df.columns if _is_altoqi_element_column(c)]
    if ignored:
        print("  Coluna técnica ignorada: " + ", ".join(str(c) for c in ignored))
        df = df.drop(columns=ignored, errors="ignore")
    return df


def _find_item_column(columns):
    """
    Analisa cada coluna e localiza a coluna identificadora do item.

    Uma coluna somente é candidata quando ``item`` aparece nos 8 primeiros
    caracteres do cabeçalho. A comparação não diferencia maiúsculas,
    minúsculas ou acentos, mas preserva a posição dos caracteres internos.

    A coluna técnica do AltoQi/Eberick nunca é considerada. Quando houver mais
    de uma candidata válida, o cabeçalho exatamente igual a ``ITEM`` tem
    prioridade; caso contrário, é usada a primeira candidata na ordem original
    da planilha.
    """
    candidates = [
        column
        for column in columns
        if (
            not _is_altoqi_element_column(column)
            and _item_in_first_eight_characters(column)
        )
    ]

    for column in candidates:
        if _column_key(column) == "item":
            return column

    return candidates[0] if candidates else None


def _clean_text_series(series):
    """Converte identificadores em texto sem transformar nulos em ``nan``."""
    return series.where(series.notna(), "").astype(str).str.strip()



def _natural_sort_key(value):
    """Ordena SEG 2 antes de SEG 10 e preserva textos comuns."""
    return [
        int(part) if part.isdigit() else part.casefold()
        for part in re.split(r"(\d+)", str(value))
    ]


def _safe_sheet_title(workbook, title):
    """Gera nome de aba válido, limitado a 31 caracteres e sem duplicidade."""
    cleaned = re.sub(r"[\[\]:*?/\\]+", " - ", str(title or "PLANILHA"))
    cleaned = re.sub(r"\s+", " ", cleaned).strip(" '")
    cleaned = cleaned or "PLANILHA"
    base = cleaned[:31]
    candidate = base
    counter = 2

    while candidate in workbook.sheetnames:
        suffix = f" ({counter})"
        candidate = f"{base[:31 - len(suffix)]}{suffix}"
        counter += 1
    return candidate


def _safe_output_filename(filename):
    """Remove caracteres inválidos do nome sugerido sem alterar a pasta."""
    cleaned = re.sub(r'[<>:"/\\|?*]+', "-", str(filename))
    cleaned = re.sub(r"\s+", " ", cleaned).strip(" .")
    return cleaned or "lista_materiais.xlsx"


def _item_in_first_eight_characters(column):
    """
    Retorna ``True`` quando ``item`` aparece nos 8 primeiros caracteres
    do cabeçalho, após remover espaços externos e ignorar a caixa.

    Exemplos aceitos:
    - ``ITEM``
    - ``Item name``
    - ``XX Item``

    Exemplo rejeitado:
    - ``AltoQi_Eberick-Item_01...``

    Nesse último caso, ``item`` aparece depois do oitavo caractere e a coluna
    pertence ao grupo de descrição/quantidade do AltoQi, não ao identificador
    da peça.
    """
    first_eight = _normalize_text(column)[:8]
    return "item" in first_eight


def _looks_like_item_header(key):
    """Aplica ao cabeçalho a regra dos 8 primeiros caracteres."""
    return _item_in_first_eight_characters(key)


def _read_excel_table(path, require_material_columns=True):
    """
    Lê a primeira tabela da primeira aba e detecta cabeçalho nas 25 primeiras
    linhas. Isso permite reabrir arquivos com título acima da tabela.
    """
    file_path = Path(path)
    if not file_path.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {file_path}")
    if file_path.suffix.lower() not in {".xlsx", ".xls"}:
        raise ValueError(f"Extensão de planilha não suportada: {file_path.suffix}")

    try:
        preview = pd.read_excel(
            file_path,
            sheet_name=0,
            header=None,
            nrows=25,
        )
    except ImportError as exc:
        raise ImportError(
            "Não foi possível ler a planilha. Para arquivos .xls, "
            "instale também o pacote xlrd."
        ) from exc
    except (OSError, ValueError) as exc:
        raise ValueError(
            f"Não foi possível ler a planilha '{file_path.name}': {exc}"
        ) from exc

    header_row = 0
    for row_index, row in preview.iterrows():
        header_values = [
            value
            for value in row.tolist()
            if _normalize_text(value)
        ]
        keys = [_column_key(value) for value in header_values]
        has_item = any(
            _looks_like_item_header(value)
            for value in header_values
        )
        has_material = any(
            (
                "descri" in key
                or "description" in key
                or "quantidade" in key
                or "quantity" in key
                or "concreto" in key
                or "forma" in key
                or re.search(r"(?:aco|o|ø)\s*\d", key)
            )
            for key in keys
        )
        if has_item and (has_material or not require_material_columns):
            header_row = int(row_index)
            break

    try:
        return pd.read_excel(
            file_path,
            sheet_name=0,
            header=header_row,
        )
    except ImportError as exc:
        raise ImportError(
            "Não foi possível ler a planilha. Para arquivos .xls, "
            "instale também o pacote xlrd."
        ) from exc
    except (OSError, ValueError) as exc:
        raise ValueError(
            f"Não foi possível carregar a tabela de '{file_path.name}': {exc}"
        ) from exc


def _qto_pair_key(column):
    key = _column_key(column)
    key = re.sub(
        r"\b(descricao|description|quantidade|quantity)\b",
        " ",
        key,
    )
    return re.sub(r"\s+", " ", key).strip()


def _pair_qto_columns(columns):
    """Pareia descrição e quantidade e rejeita exportações incompletas."""
    description_columns = [
        column
        for column in columns
        if (
            "descri" in _normalize_text(column)
            or "description" in _normalize_text(column)
        )
    ]
    quantity_columns = [
        column
        for column in columns
        if (
            "quantidade" in _normalize_text(column)
            or "quantity" in _normalize_text(column)
        )
    ]

    if not description_columns and not quantity_columns:
        return []

    if not description_columns or not quantity_columns:
        raise ValueError(
            "A planilha possui somente um dos grupos necessários: "
            "Descrição ou Quantidade. Verifique a exportação do AltoQi/Eberick."
        )

    if len(description_columns) != len(quantity_columns):
        raise ValueError(
            "Quantidade incompatível de colunas: "
            f"{len(description_columns)} descrição(ões) e "
            f"{len(quantity_columns)} quantidade(s)."
        )

    descriptions_by_key = {}
    quantities_by_key = {}
    for column in description_columns:
        descriptions_by_key.setdefault(_qto_pair_key(column), []).append(column)
    for column in quantity_columns:
        quantities_by_key.setdefault(_qto_pair_key(column), []).append(column)

    if set(descriptions_by_key) == set(quantities_by_key):
        pairs = []
        for key in sorted(descriptions_by_key, key=_natural_sort_key):
            descriptions = sorted(
                descriptions_by_key[key],
                key=_natural_sort_key,
            )
            quantities = sorted(
                quantities_by_key[key],
                key=_natural_sort_key,
            )
            if len(descriptions) != len(quantities):
                break
            pairs.extend(zip(descriptions, quantities))
        else:
            return pairs

    descriptions = sorted(description_columns, key=_natural_sort_key)
    quantities = sorted(quantity_columns, key=_natural_sort_key)
    return list(zip(descriptions, quantities))


def _drop_generated_total_rows(dataframe):
    """
    Remove linhas auxiliares geradas pelo próprio programa para evitar
    duplicação de quantitativos ao unir planilhas já processadas.
    """
    if "ITEM" not in dataframe.columns or dataframe.empty:
        return dataframe

    item_key = dataframe["ITEM"].map(_normalize_text)
    remove_mask = item_key.isin({"total", "total geral"})
    remove_mask |= item_key.str.startswith("subtotal ", na=False)

    non_numeric_columns = [
        column
        for column in dataframe.columns
        if column not in {
            "ITEM",
            "CONCRETO_m3",
            "FORMA_m2",
        }
        and not str(column).startswith("ACO_")
    ]
    for column in non_numeric_columns:
        values = dataframe[column].map(_normalize_text)
        remove_mask |= values.isin({"total", "total geral"})
        remove_mask |= values.str.startswith("subtotal ", na=False)

    numeric_columns = [
        column
        for column in dataframe.columns
        if column in {"CONCRETO_m3", "FORMA_m2"}
        or str(column).startswith("ACO_")
    ]
    blank_item = item_key.eq("")
    if numeric_columns:
        has_quantity = (
            dataframe[numeric_columns]
            .fillna(0)
            .abs()
            .sum(axis=1)
            .gt(0)
        )
        remove_mask |= blank_item & ~has_quantity
    else:
        remove_mask |= blank_item

    removed = int(remove_mask.sum())
    if removed:
        print(f"  Linhas auxiliares ignoradas: {removed}")
    return dataframe.loc[~remove_mask].reset_index(drop=True)


def _series_sum(dataframe, column):
    if column not in dataframe.columns:
        return 0.0
    return float(pd.to_numeric(dataframe[column], errors="coerce").fillna(0).sum())



def _save_workbook(workbook, file_path):
    """Salva a saída com mensagem clara para arquivo bloqueado/permissão."""
    destination = Path(file_path)
    try:
        workbook.save(destination)
    except PermissionError as exc:
        raise PermissionError(
            f"Não foi possível salvar '{destination.name}'. "
            "Feche o arquivo no Excel e verifique a permissão da pasta."
        ) from exc
    except OSError as exc:
        raise OSError(
            f"Falha ao salvar a planilha em '{destination}': {exc}"
        ) from exc


def _copy_output_to_cache(file_path):
    """Preserva a cópia de cache existente, somente quando o drive D: existe."""
    if os.name != "nt":
        return
    cache_root = Path("D:/")
    if not cache_root.exists():
        print("  (cópia de cache ignorada: drive D: não disponível)")
        return

    cache = Path(r"D:\Cache\Claude")
    try:
        cache.mkdir(parents=True, exist_ok=True)
        destination = cache / Path(file_path).name
        shutil.copy2(file_path, destination)
        print(f"\nCopiado para: {destination}")
    except OSError as exc:
        print(f"  (cópia de cache falhou: {exc})")


def _open_output_file(file_path):
    if os.name != "nt" or not hasattr(os, "startfile"):
        print("  Abertura automática disponível somente no Windows.")
        return
    try:
        os.startfile(file_path)
        print(f"Abrindo: {file_path}")
    except OSError as exc:
        print(f"  (erro ao abrir o arquivo: {exc})")


def _show_unexpected_error(exc):
    """Exibe o erro sem ocultar o traceback técnico no console."""
    traceback.print_exc()
    root = None
    try:
        root = _root()
        messagebox.showerror(
            "Erro no processamento",
            f"{type(exc).__name__}: {exc}\n\n"
            "Nenhum arquivo de saída foi considerado concluído.",
            parent=root,
        )
    except tk.TclError:
        print(f"ERRO: {type(exc).__name__}: {exc}")
    finally:
        if root is not None:
            try:
                root.destroy()
            except tk.TclError:
                pass


def _ask_choice(title, prompt, options):
    """Retorna a opção escolhida sem lançar exceção dentro do callback."""
    if not options:
        raise ValueError("A lista de opções não pode estar vazia.")

    root = _root()
    result = [options[0]]
    cancelled = [False]
    try:
        var = tk.StringVar(value=options[0])
        win = tk.Toplevel(root)
        win.title(title)
        win.attributes("-topmost", True)
        tk.Label(
            win,
            text=prompt,
            font=("Arial", 10),
            padx=10,
            pady=8,
        ).pack()
        for option in options:
            tk.Radiobutton(
                win,
                text=option,
                variable=var,
                value=option,
                font=("Arial", 10),
                anchor="w",
            ).pack(fill="x", padx=20)

        def accept():
            result[0] = var.get()
            win.destroy()

        def cancel():
            if _confirm_cancel(win):
                cancelled[0] = True
                win.destroy()

        tk.Button(win, text="OK", command=accept, width=10).pack(pady=8)
        tk.Button(
            win,
            text="Cancelar",
            command=cancel,
            fg="white",
            bg="red",
            width=10,
        ).pack(pady=2)
        win.protocol("WM_DELETE_WINDOW", cancel)
        win.wait_window()

        if cancelled[0]:
            raise SystemExit("Cancelado pelo usuário")
        return result[0]
    finally:
        try:
            root.destroy()
        except tk.TclError:
            pass

def _ask_yes_no(title, prompt):
    """Solicita SIM/NÃO e trata cancelamento fora dos callbacks Tk."""
    root = _root()
    result = [False]
    cancelled = [False]
    try:
        win = tk.Toplevel(root)
        win.title(title)
        win.attributes("-topmost", True)
        tk.Label(
            win,
            text=prompt,
            font=("Arial", 10),
            padx=10,
            pady=8,
        ).pack()

        def finish(value):
            result[0] = value
            win.destroy()

        def cancel():
            if _confirm_cancel(win):
                cancelled[0] = True
                win.destroy()

        frame = tk.Frame(win)
        frame.pack(pady=8)
        tk.Button(
            frame,
            text="SIM",
            command=lambda: finish(True),
            width=10,
        ).pack(side="left", padx=8)
        tk.Button(
            frame,
            text="NÃO",
            command=lambda: finish(False),
            width=10,
        ).pack(side="left", padx=8)
        tk.Button(
            frame,
            text="Cancelar",
            command=cancel,
            fg="white",
            bg="red",
            width=10,
        ).pack(side="left", padx=8)
        win.protocol("WM_DELETE_WINDOW", cancel)
        win.wait_window()

        if cancelled[0]:
            raise SystemExit("Cancelado pelo usuário")
        return result[0]
    finally:
        try:
            root.destroy()
        except tk.TclError:
            pass


def _format_merge_title(meta, section):
    tipo_part = f" - {meta['tipo']}" if meta.get("tipo") else ""
    base = f"UNIAO DE LISTAS - {meta['elemento']} - {meta['pavimento']} - {meta['obra']}{tipo_part} - {meta['data']}"
    return f"{base} - {section}" if section else base


def _write_merge_title(ws, title, end_col):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=end_col)
    ws["A1"].value = title
    ws["A1"].font = Font(name="Arial", bold=True, size=12, color="1F4E79")
    ws["A1"].alignment = C
    ws.row_dimensions[1].height = 22


def _merge_header_row(ws, meta, section, end_col):
    if not meta:
        return 1
    _write_merge_title(ws, _format_merge_title(meta, section), end_col)
    return 3


def _build_merge_sheet(workbook, title, dataframe, bitolas, meta=None):
    sheet = workbook.create_sheet(_safe_sheet_title(workbook, title))
    field_names = [
        "ITEM",
        "NIVEL",
        "CONCRETO_m3",
        "FORMA_m2",
    ] + [f"ACO_{bitola}mm_kg" for bitola in bitolas]
    headers = [
        "ITEM",
        "NIVEL",
        "CONCRETO_m3",
        "FORMA_m2",
    ] + [f"O{bitola}mm kg" for bitola in bitolas]

    header_row = _merge_header_row(
        sheet,
        meta,
        title,
        len(headers),
    )
    for column_index, header in enumerate(headers, 1):
        _h(sheet, header_row, column_index, header, fl=H_PILAR)

    records = dataframe.to_dict(orient="records")
    for row_index, row in enumerate(
        records,
        start=header_row + 1,
    ):
        for column_index, field in enumerate(field_names, 1):
            is_numeric = (
                field.startswith("ACO_")
                or field in {"CONCRETO_m3", "FORMA_m2"}
            )
            value = row.get(field, 0.0 if is_numeric else "")
            if is_numeric and (value is None or pd.isna(value)):
                value = 0.0
            _c(
                sheet,
                row_index,
                column_index,
                value,
                fl=ALT[(row_index - (header_row + 1)) % 2],
                al=R if is_numeric else L,
                fmt="#,##0.000" if is_numeric else None,
            )

    widths = [30, 18, 14, 14] + [14] * len(bitolas)
    for column_index, width in enumerate(widths, 1):
        sheet.column_dimensions[get_column_letter(column_index)].width = width
    sheet.freeze_panes = f"A{header_row + 1}"
    return sheet


def _build_merge_summary_sheet(workbook, summaries, bitolas, meta=None):
    sheet = workbook.create_sheet(
        _safe_sheet_title(workbook, "RESUMO POR SEGMENTO")
    )
    headers = [
        "ITEM",
        "SEGMENTO",
        "CONCRETO_m3",
        "FORMA_m2",
    ] + [f"O{bitola}mm kg" for bitola in bitolas]
    header_row = _merge_header_row(
        sheet,
        meta,
        "RESUMO POR SEGMENTO",
        len(headers),
    )
    piece_label = (
        (meta or {}).get("elemento")
        or summaries[0]["group_label"]
    )

    for column_index, header in enumerate(headers, 1):
        _h(sheet, header_row, column_index, header, fl=H_PILAR)

    for row_index, summary in enumerate(
        summaries,
        start=header_row + 1,
    ):
        values = [
            piece_label,
            summary["segment"],
            summary["totals"]["CONCRETO_m3"],
            summary["totals"]["FORMA_m2"],
        ]
        values.extend(
            summary["totals"].get(
                f"ACO_{bitola}mm_kg",
                0.0,
            )
            for bitola in bitolas
        )

        for column_index, value in enumerate(values, 1):
            is_numeric = column_index >= 3
            _c(
                sheet,
                row_index,
                column_index,
                value,
                fl=ALT[(row_index - (header_row + 1)) % 2],
                al=R if is_numeric else L,
                fmt="#,##0.000" if is_numeric else None,
            )

    totals = {
        "CONCRETO_m3": sum(
            summary["totals"]["CONCRETO_m3"]
            for summary in summaries
        ),
        "FORMA_m2": sum(
            summary["totals"]["FORMA_m2"]
            for summary in summaries
        ),
    }
    for bitola in bitolas:
        totals[f"ACO_{bitola}mm_kg"] = sum(
            summary["totals"].get(
                f"ACO_{bitola}mm_kg",
                0.0,
            )
            for summary in summaries
        )

    total_row = header_row + len(summaries) + 1
    _c(sheet, total_row, 1, piece_label, fnt=WF, fl=H_TOTAL, al=L)
    _c(sheet, total_row, 2, "TOTAL GERAL", fnt=WF, fl=H_TOTAL, al=L)
    _c(
        sheet,
        total_row,
        3,
        totals["CONCRETO_m3"],
        fnt=WF,
        fl=H_TOTAL,
        al=R,
        fmt="#,##0.000",
    )
    _c(
        sheet,
        total_row,
        4,
        totals["FORMA_m2"],
        fnt=WF,
        fl=H_TOTAL,
        al=R,
        fmt="#,##0.000",
    )
    for column_index, bitola in enumerate(bitolas, start=5):
        _c(
            sheet,
            total_row,
            column_index,
            totals.get(f"ACO_{bitola}mm_kg", 0.0),
            fnt=WF,
            fl=H_TOTAL,
            al=R,
            fmt="#,##0.000",
        )

    widths = [24, 20, 14, 14] + [14] * len(bitolas)
    for column_index, width in enumerate(widths, 1):
        sheet.column_dimensions[get_column_letter(column_index)].width = width
    sheet.freeze_panes = f"A{header_row + 1}"
    return sheet

def _build_merge_total_sheet(
    workbook,
    total,
    piece_label,
    bitolas,
    meta=None,
):
    title = f"RESUMO TOTAL {piece_label}"
    sheet = workbook.create_sheet(
        _safe_sheet_title(workbook, title)
    )
    headers = [
        "ITEM",
        "CONCRETO_m3",
        "FORMA_m2",
    ] + [f"O{bitola}mm kg" for bitola in bitolas]
    header_row = _merge_header_row(
        sheet,
        meta,
        title,
        len(headers),
    )
    for column_index, header in enumerate(headers, 1):
        _h(sheet, header_row, column_index, header, fl=H_PILAR)

    values = [
        piece_label,
        total["CONCRETO_m3"],
        total["FORMA_m2"],
    ] + [
        total.get(f"ACO_{bitola}mm_kg", 0.0)
        for bitola in bitolas
    ]
    for column_index, value in enumerate(values, 1):
        is_numeric = column_index >= 2
        _c(
            sheet,
            header_row + 1,
            column_index,
            value,
            fnt=WF,
            fl=H_TOTAL,
            al=R if is_numeric else L,
            fmt="#,##0.000" if is_numeric else None,
        )

    widths = [30, 16, 16] + [14] * len(bitolas)
    for column_index, width in enumerate(widths, 1):
        sheet.column_dimensions[get_column_letter(column_index)].width = width
    sheet.freeze_panes = f"A{header_row + 1}"
    return sheet


def _analyze_merge_files(paths):
    if not paths:
        raise ValueError("Nenhuma planilha foi selecionada para união.")

    summaries = []
    for index, path in enumerate(paths, start=1):
        print(f"  Lendo: {Path(path).name}")
        dataframe, bitolas, _ = load_qto(path)

        if dataframe.empty:
            raise ValueError(
                f"A planilha '{Path(path).name}' não possui itens válidos."
            )

        categories = sorted(
            {
                category
                for category in (
                    _extract_element_category(value)
                    for value in dataframe["ITEM"]
                )
                if category
            },
            key=_natural_sort_key,
        )
        if not categories:
            categories = [
                _extract_element_category(Path(path).stem)
                or "PEÇAS"
            ]

        levels = sorted(
            {
                _normalize_text(value)
                for value in dataframe["NIVEL"]
                if _normalize_text(value)
            },
            key=_natural_sort_key,
        )
        segment = _extract_segment_label(Path(path).name, index)

        totals = {
            "CONCRETO_m3": _series_sum(dataframe, "CONCRETO_m3"),
            "FORMA_m2": _series_sum(dataframe, "FORMA_m2"),
        }
        for bitola in bitolas:
            totals[f"ACO_{bitola}mm_kg"] = _series_sum(
                dataframe,
                f"ACO_{bitola}mm_kg",
            )

        summaries.append(
            {
                "path": Path(path),
                "df": dataframe,
                "bitolas": bitolas,
                "categories": categories,
                "levels": levels,
                "segment": segment,
                "group_label": (
                    categories[0].upper()
                    if categories
                    else "PEÇAS"
                ),
                "totals": totals,
            }
        )

    group_sets = [set(summary["categories"]) for summary in summaries]
    if len({tuple(sorted(group)) for group in group_sets}) > 1:
        details = "\n".join(
            f"{summary['path'].name}: "
            f"{', '.join(summary['categories'])}"
            for summary in summaries
        )
        root = _root()
        try:
            messagebox.showinfo(
                "Análise de itens",
                f"Diferenças de itens encontradas:\n\n{details}",
                parent=root,
            )
        finally:
            root.destroy()
        if not _ask_yes_no(
            "Divergência de itens",
            "Quer continuar mesmo assim?",
        ):
            raise SystemExit("Operação cancelada pelo usuário")

    level_sets = [set(summary["levels"]) for summary in summaries]
    if len({tuple(sorted(levels)) for levels in level_sets}) > 1:
        details = "\n".join(
            f"{summary['path'].name}: "
            f"{', '.join(summary['levels'])}"
            for summary in summaries
        )
        root = _root()
        try:
            messagebox.showinfo(
                "Análise de níveis/pavimentos",
                f"Diferenças de nível/pavimento encontradas:\n\n{details}",
                parent=root,
            )
        finally:
            root.destroy()
        if not _ask_yes_no(
            "Divergência de níveis",
            "Quer continuar mesmo assim?",
        ):
            raise SystemExit("Operação cancelada pelo usuário")

    return summaries


def merge_material_lists():
    print("\n=== UNIR LISTAS DE MATERIAIS ===")
    meta = coletar_metadados_uniao()
    paths = _ask_files("Selecione as planilhas para unir")
    summaries = _analyze_merge_files(paths)

    initial_name = _safe_output_filename(
        f"{meta['nome_base']}.xlsx"
    )
    file_out = _ask_save(
        "Salvar nova planilha unida como...",
        initial_name,
    )

    workbook = Workbook()
    if workbook.active is not None:
        workbook.remove(workbook.active)

    all_bitolas = sorted(
        {
            bitola
            for summary in summaries
            for bitola in summary["bitolas"]
        }
    )
    source_group_label = summaries[0]["group_label"]
    final_group_label = meta["elemento"] or source_group_label

    for summary in summaries:
        dataframe = summary["df"].copy()
        for bitola in all_bitolas:
            column = f"ACO_{bitola}mm_kg"
            if column not in dataframe.columns:
                dataframe[column] = 0.0

        source_label = final_group_label
        title = f"TODOS - {source_label} {summary['segment']}"
        auxiliary_meta = dict(meta)
        auxiliary_meta["elemento"] = source_label
        _build_merge_sheet(
            workbook,
            title,
            dataframe,
            all_bitolas,
            auxiliary_meta,
        )

    _build_merge_summary_sheet(
        workbook,
        summaries,
        all_bitolas,
        dict(meta),
    )

    total = {
        "CONCRETO_m3": sum(
            summary["totals"]["CONCRETO_m3"]
            for summary in summaries
        ),
        "FORMA_m2": sum(
            summary["totals"]["FORMA_m2"]
            for summary in summaries
        ),
    }
    for bitola in all_bitolas:
        total[f"ACO_{bitola}mm_kg"] = sum(
            summary["totals"].get(
                f"ACO_{bitola}mm_kg",
                0.0,
            )
            for summary in summaries
        )

    _build_merge_total_sheet(
        workbook,
        total,
        final_group_label,
        all_bitolas,
        meta,
    )
    _save_workbook(workbook, file_out)
    _copy_output_to_cache(file_out)

    if _ask_yes_no(
        "Planilha gerada",
        f"Planilha gerada:\n{file_out.name}\n\n"
        "Deseja abrir a planilha agora?",
    ):
        _open_output_file(file_out)

    print(f"\nSalvo em:\n{file_out}")

# ─── METADADOS ───────────────────────────────────────────────────────────────

def coletar_metadados_uniao():
    print("\n=== UNIR LISTAS DE MATERIAIS ===")
    print("Configurando metadados da união...\n")

    def env_or_ask(env_name, title, prompt, default=""):
        value = os.environ.get(env_name, "").strip().upper()
        if value:
            return value
        return (_ask_str(title, prompt, default) or default).strip().upper()

    piece = env_or_ask(
        "LM_ELEMENTO",
        "Peça",
        "Qual peça é?\n(ex.: LAJES, VIGAS, PILARES, SAPATAS)",
        "PEÇA",
    )
    floor = env_or_ask(
        "LM_PAVIMENTO",
        "Pavimento",
        "Qual PAVIMENTO?\n(ex.: PAV BALDRAMES, TÉRREO, 1 PAV, 2 PAV)",
        "PAV",
    )
    project = env_or_ask(
        "LM_OBRA",
        "Obra",
        "Qual OBRA?\n(ex.: IFMTR, IFMTC, OBRA XYZ)",
        "IFMT",
    )

    project_type_env = os.environ.get("LM_TIPO", "").strip().upper()
    project_type = (
        project_type_env
        if project_type_env in {"EXECUTIVO", "OFICIAL"}
        else _ask_choice(
            "Tipo",
            "Projeto EXECUTIVO ou OFICIAL?",
            ["EXECUTIVO", "OFICIAL"],
        )
    )

    date_env = os.environ.get("LM_DATA", "").strip()
    expedition_date = date_env or _ask_str(
        "Data",
        "Data de expedição?\n(deixar vazio = hoje)",
        default=datetime.today().strftime("%d.%m.%Y"),
    )
    expedition_date = (
        expedition_date
        or datetime.today().strftime("%d.%m.%Y")
    )

    base_name = (
        f"UNIAO LISTAS - {piece} - {floor} - {project} - "
        f"{project_type} - {expedition_date}"
    )

    print(f"  Peça      : {piece}")
    print(f"  Pavimento : {floor}")
    print(f"  Obra      : {project}")
    print(f"  Tipo      : {project_type}")
    print(f"  Data      : {expedition_date}")
    print(f"  Arquivo   : {_safe_output_filename(base_name + '.xlsx')}\n")

    return {
        "elemento": piece,
        "pavimento": floor,
        "obra": project,
        "tipo": project_type,
        "data": expedition_date,
        "nome_base": base_name,
    }


def coletar_metadados():
    print("\n=== LISTA DE MATERIAIS IFMT ===")
    print("Configurando metadados...\n")

    def env_or_ask(env_name, title, prompt, default=""):
        value = os.environ.get(env_name, "").strip().upper()
        if value:
            return value
        return (_ask_str(title, prompt, default) or default).strip().upper()

    piece = env_or_ask(
        "LM_ELEMENTO",
        "Peça",
        "Qual peça é?\n(ex.: LAJES, VIGAS, PILARES, SAPATAS)",
        "PEÇA",
    )
    floor = env_or_ask(
        "LM_PAVIMENTO",
        "Pavimento",
        "Qual PAVIMENTO?\n(ex.: PAV BALDRAMES, TÉRREO, 1 PAV, 2 PAV)",
        "PAV",
    )
    segment = env_or_ask(
        "LM_SEGMENTO",
        "Segmento",
        "Qual SEGMENTO?\n(ex.: SEG 1, SEG 2, SEG 3, TODOS)",
        "TODOS",
    )
    project = env_or_ask(
        "LM_OBRA",
        "Obra",
        "Qual OBRA?\n(ex.: IFMTR, IFMTC, OBRA XYZ)",
        "IFMT",
    )

    project_type_env = os.environ.get("LM_TIPO", "").strip().upper()
    project_type = (
        project_type_env
        if project_type_env in {"EXECUTIVO", "OFICIAL"}
        else _ask_choice(
            "Tipo",
            "Projeto EXECUTIVO ou OFICIAL?",
            ["EXECUTIVO", "OFICIAL"],
        )
    )

    date_env = os.environ.get("LM_DATA", "").strip()
    expedition_date = date_env or _ask_str(
        "Data",
        "Data de expedição?\n(deixar vazio = hoje)",
        default=datetime.today().strftime("%d.%m.%Y"),
    )
    expedition_date = (
        expedition_date
        or datetime.today().strftime("%d.%m.%Y")
    )

    mode_env = os.environ.get("LM_MODO", "").strip().upper()
    if mode_env in {"A", "RESUMO"}:
        summary_mode = True
    elif mode_env in {"B", "COMPLETO"}:
        summary_mode = False
    else:
        selected_mode = _ask_choice(
            "Modo de saída",
            "O que você quer gerar?",
            [
                "(A) Só RESUMO (totais por material)",
                "(B) LISTA COMPLETA (todos os itens + resumo)",
            ],
        )
        summary_mode = "(A)" in selected_mode

    observation_env = os.environ.get("LM_OBS", "").strip().upper()
    observation = observation_env or _ask_str(
        "Observação",
        "Observação opcional (ex.: SEM ESCADAS, BLOCO A)\n"
        "Deixe vazio para nenhuma.",
        default="",
    ).strip().upper()

    base_name = (
        f"{piece} - {floor} - {project} - {project_type} - "
        f"{segment} - {expedition_date}"
    )
    if observation:
        base_name += f" - {observation}"

    print(f"  Peça      : {piece}")
    print(f"  Pavimento : {floor}")
    print(f"  Segmento  : {segment}")
    print(f"  Obra      : {project}")
    print(f"  Tipo      : {project_type}")
    print(f"  Data      : {expedition_date}")
    print(f"  Observação: {observation or '(nenhuma)'}")
    print(f"  Modo      : {'RESUMO' if summary_mode else 'COMPLETO'}")
    print(f"  Arquivo   : {_safe_output_filename(base_name + '.xlsx')}\n")

    return {
        "elemento": piece,
        "pavimento": floor,
        "segmento": segment,
        "obra": project,
        "tipo": project_type,
        "data": expedition_date,
        "observacao": observation,
        "nome_base": base_name,
        "modo_resumo": summary_mode,
    }

def coletar_arquivos(meta):
    print("Selecione os arquivos de entrada:\n")
    print("1. Planilha TOTAL (todos os itens)...")
    file_qto = _ask_file("1 - Planilha TOTAL (todos os itens)")
    print(f"   OK: {file_qto.name}")

    segment_files = {}
    has_segments = _ask_yes_no(
        "Segmentos",
        "Você tem arquivos separados por segmento?\n"
        "(SEG 1, SEG 2 etc.)",
    )
    if has_segments:
        index = 1
        while _ask_yes_no(
            "Segmentos",
            f"Adicionar arquivo SEG {index}?",
        ):
            print(f"{index + 1}. Planilha SEG {index}...")
            file_path = _ask_file(f"Planilha SEG {index}")
            segment_files[f"SEG {index}"] = file_path
            print(f"   OK: {file_path.name}")
            index += 1

    initial_name = _safe_output_filename(
        f"{meta['nome_base']}.xlsx"
    )
    file_out = _ask_save(
        "Salvar planilha como...",
        initial_name,
    )
    print(f"\nArquivo de saída: {file_out.name}\n")
    return file_qto, segment_files, file_out

# ─── PARSE ───────────────────────────────────────────────────────────────────

NIVEL_MAP = {
    "TERREO":"TERREO","TÉRREO":"TERREO","Térreo":"TERREO","Terreo":"TERREO",
    "1 PAVIMENTO":"TERREO","1 Pavimento":"TERREO","Pavimento 1":"TERREO","PAVIMENTO 1":"TERREO",
    "BALDRAME":"BALDRAME","Baldrame":"BALDRAME","PAV BALDRAMES":"BALDRAME","BALDRAMES":"BALDRAME",
    "2 PAVIMENTO":"PAV 2","2 Pavimento":"PAV 2","Pavimento 2":"PAV 2","2° PAV":"PAV 2",
    "3 PAVIMENTO":"PAV 3","3 Pavimento":"PAV 3","Pavimento 3":"PAV 3","3° PAV":"PAV 3",
}

def normalizar_nivel(value, default="SEM NIVEL"):
    """Normaliza níveis conhecidos e trata células vazias de forma explícita."""
    if not _normalize_text(value):
        return str(default or "SEM NIVEL").strip()

    original = str(value).strip()
    normalized = _normalize_text(original)
    for source_name, target_name in NIVEL_MAP.items():
        if _normalize_text(source_name) == normalized:
            return target_name
    return original

def parse_qty(value):
    """
    Converte quantidades com decimal brasileiro ou internacional.

    Unidades são removidas antes da leitura; assim, ``1,5 m3`` resulta em 1.5
    e uma célula contendo apenas ``m3`` resulta em zero.
    """
    if value is None:
        return 0.0
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)) and not pd.isna(value):
        return float(value)
    if pd.isna(value):
        return 0.0

    text = str(value).strip()
    if not text:
        return 0.0

    text = re.sub(
        r"(?i)(?<![a-z])(?:kg|mm|cm|m(?:2|3|²|³)?)(?![a-z])",
        " ",
        text,
    )
    match = re.search(r"[+-]?\d(?:[\d\s.,]*\d)?", text)
    if not match:
        return 0.0

    token = re.sub(r"\s+", "", match.group(0))
    sign = ""
    if token.startswith(("+", "-")):
        sign, token = token[0], token[1:]

    if "," in token and "." in token:
        decimal_separator = "," if token.rfind(",") > token.rfind(".") else "."
        thousands_separator = "." if decimal_separator == "," else ","
        token = token.replace(thousands_separator, "")
        token = token.replace(decimal_separator, ".")
    elif "," in token:
        parts = token.split(",")
        if len(parts) > 2 and all(len(part) == 3 for part in parts[1:]):
            token = "".join(parts)
        else:
            token = "".join(parts[:-1]) + "." + parts[-1]
    elif token.count(".") > 1:
        parts = token.split(".")
        if all(len(part) == 3 for part in parts[1:]):
            token = "".join(parts)
        else:
            token = "".join(parts[:-1]) + "." + parts[-1]

    try:
        return float(sign + token)
    except ValueError:
        return 0.0

def extract_bitola(description):
    match = re.search(
        r"[Øø]?\s*([\d,.]+)\s*mm",
        str(description),
        re.IGNORECASE,
    )
    return float(match.group(1).replace(",", ".")) if match else None


def _bitola_from_aco_column(column):
    match = re.search(
        r"ACO_([\d.,]+)mm",
        str(column),
        re.IGNORECASE,
    )
    return float(match.group(1).replace(",", ".")) if match else None


def _norm_cols(dataframe):
    """
    Limpa quebras de linha e caracteres de exportação dos cabeçalhos.

    Os nomes originais dos dados são preservados; somente os cabeçalhos são
    ajustados para permitir comparação consistente.
    """
    rename = {}
    for column in dataframe.columns:
        normalized = str(column)
        normalized = re.sub(r"x000D_", "", normalized, flags=re.IGNORECASE)
        normalized = re.sub(r"[\r\n]+", " ", normalized)
        normalized = re.sub(
            r"Forma\s*-\s*Estrutura\s*-\s*Concreto",
            "Forma",
            normalized,
            flags=re.IGNORECASE,
        )
        normalized = re.sub(
            r"\b(?:m³|m3|m2|kg)\b",
            "",
            normalized,
            flags=re.IGNORECASE,
        )
        normalized = re.sub(r"\s{2,}", " ", normalized).strip()
        if normalized != str(column):
            rename[column] = normalized
    return dataframe.rename(columns=rename) if rename else dataframe

def normalize_desc(description):
    if description is None or pd.isna(description):
        return ""
    text = str(description).strip()
    if not text:
        return ""

    text = re.sub(
        r"Forma\s*-\s*Estrutura\s*-\s*Concreto",
        "Forma",
        text,
        flags=re.IGNORECASE,
    )
    text = re.sub(
        r"Forma\s*-\s*Estrutura",
        "Forma",
        text,
        flags=re.IGNORECASE,
    )
    text = re.sub(
        r"\b(?:m³|m3|m2|kg)\b",
        "",
        text,
        flags=re.IGNORECASE,
    )
    match = re.match(
        r"^(Concreto\s*-\s*C-\d+(?:[.,]\d+)?)(?:\s*[-–—].*)?$",
        text,
        flags=re.IGNORECASE,
    )
    if match:
        text = match.group(1)
    return re.sub(r"\s{2,}", " ", text).strip()

def load_qto(path, default_nivel=None):
    """
    Retorna ``(dataframe, bitolas, possui_coluna_nivel)``.

    Regras principais:
    - remove ``AltoQi_Eberick_Elemento / Elemento``;
    - usa a coluna em que ``item`` aparece nos 8 primeiros caracteres como origem de ``ITEM``;
    - aceita listas já processadas sem duplicar linhas TOTAL/SUBTOTAL;
    - preserva ``ITEM`` como primeira coluna.
    """
    dataframe = _read_excel_table(path, require_material_columns=True)
    dataframe = _norm_cols(dataframe)
    dataframe = _drop_ignored_source_columns(dataframe)
    dataframe = dataframe.dropna(axis=1, how="all")

    columns = [str(column).strip() for column in dataframe.columns]
    qto_pairs = _pair_qto_columns(columns)

    if not qto_pairs:
        item_column = _find_item_column(columns)
        if item_column is None:
            raise ValueError(
                "Nenhuma coluna com 'item' nos 8 primeiros caracteres do título foi encontrada."
            )

        if item_column != "ITEM":
            dataframe = dataframe.rename(columns={item_column: "ITEM"})
        dataframe["ITEM"] = _clean_text_series(dataframe["ITEM"])

        columns = [str(column).strip() for column in dataframe.columns]
        level_column = _find_column(
            columns,
            ["nivel", "nível", "pavimento", "pav", "piso", "layer"],
            exclude=["name", "nome"],
        )
        has_level_column = level_column is not None
        if level_column is not None and level_column != "NIVEL":
            dataframe = dataframe.rename(columns={level_column: "NIVEL"})
        elif level_column is None:
            dataframe["NIVEL"] = default_nivel or "UNICO"

        columns = [str(column).strip() for column in dataframe.columns]
        upper_lookup = {str(column).upper(): column for column in dataframe.columns}

        concrete_column = upper_lookup.get("CONCRETO_M3")
        if concrete_column is None:
            concrete_column = _find_column(
                columns,
                ["concreto"],
                exclude=["kg", "mm", "descri"],
            )

        form_column = upper_lookup.get("FORMA_M2")
        if form_column is None:
            form_column = _find_column(
                columns,
                ["forma"],
                exclude=["infor", "descri"],
            )

        steel_sources = {}
        for column in list(dataframe.columns):
            bitola = _bitola_from_aco_column(column)
            if bitola is None:
                match = re.match(
                    r"[OØø]\s*([\d.,]+)\s*mm",
                    str(column),
                    flags=re.IGNORECASE,
                )
                if match:
                    bitola = float(match.group(1).replace(",", "."))
            if bitola is not None:
                canonical = f"ACO_{bitola}mm_kg"
                steel_sources.setdefault(canonical, []).append(column)

        has_material_column = bool(
            concrete_column is not None
            or form_column is not None
            or steel_sources
        )
        if not has_material_column:
            raise ValueError(
                "A tabela possui ITEM, mas não contém colunas reconhecidas "
                "de CONCRETO, FORMA ou AÇO."
            )

        if concrete_column is not None and concrete_column != "CONCRETO_m3":
            dataframe = dataframe.rename(
                columns={concrete_column: "CONCRETO_m3"}
            )
        if form_column is not None and form_column != "FORMA_m2":
            dataframe = dataframe.rename(columns={form_column: "FORMA_m2"})

        for canonical, source_columns in steel_sources.items():
            values = pd.Series(0.0, index=dataframe.index)
            for source_column in source_columns:
                values = values.add(
                    dataframe[source_column].apply(parse_qty),
                    fill_value=0,
                )
            dataframe[canonical] = values

        removable_steel_sources = [
            source_column
            for canonical, source_columns in steel_sources.items()
            for source_column in source_columns
            if source_column != canonical
        ]
        if removable_steel_sources:
            dataframe = dataframe.drop(
                columns=removable_steel_sources,
                errors="ignore",
            )

        for required in ("CONCRETO_m3", "FORMA_m2"):
            if required not in dataframe.columns:
                dataframe[required] = 0.0

        steel_columns = sorted(
            {
                column
                for column in dataframe.columns
                if _bitola_from_aco_column(column) is not None
            },
            key=lambda column: _bitola_from_aco_column(column) or 0.0,
        )
        numeric_columns = ["CONCRETO_m3", "FORMA_m2"] + steel_columns
        for column in numeric_columns:
            dataframe[column] = dataframe[column].apply(parse_qty).fillna(0.0)

        dataframe["NIVEL"] = dataframe["NIVEL"].apply(
            lambda value: normalizar_nivel(
                value,
                default=default_nivel or "UNICO",
            )
        )
        dataframe = _drop_generated_total_rows(dataframe)

        blank_item = dataframe["ITEM"].map(_normalize_text).eq("")
        if blank_item.any():
            raise ValueError(
                f"Foram encontradas {int(blank_item.sum())} linha(s) com "
                "quantitativos, mas sem ITEM."
            )

        bitolas = sorted(
            {
                bitola
                for bitola in (
                    _bitola_from_aco_column(column)
                    for column in steel_columns
                )
                if bitola is not None
            }
        )
        ordered = [
            "ITEM",
            "NIVEL",
            "CONCRETO_m3",
            "FORMA_m2",
        ] + [
            f"ACO_{bitola}mm_kg"
            for bitola in bitolas
            if f"ACO_{bitola}mm_kg" in dataframe.columns
        ]
        remaining = [
            column
            for column in dataframe.columns
            if column not in ordered
        ]
        return dataframe[ordered + remaining], bitolas, has_level_column

    item_column = _find_item_column(columns)
    if item_column is None:
        raise ValueError(
            "Nenhuma coluna com 'item' nos 8 primeiros caracteres do título foi encontrada. A coluna técnica "
            "'AltoQi_Eberick_Elemento / Elemento' é ignorada e não pode "
            "ser usada como nome do item."
        )

    if item_column != "ITEM":
        dataframe = dataframe.rename(columns={item_column: "ITEM"})
    dataframe["ITEM"] = _clean_text_series(dataframe["ITEM"])

    columns = [str(column).strip() for column in dataframe.columns]
    level_column = _find_column(
        columns,
        ["layer", "nivel", "nível", "pavimento", "pav", "piso"],
        exclude=["name", "nome"],
    )
    has_level_column = level_column is not None
    if not has_level_column:
        print(
            f"  Sem coluna Layer/Nível -> nível = "
            f"'{default_nivel or 'UNICO'}'"
        )

    print(
        f"  item_col='{item_column}'  "
        f"col_layer='{level_column or '(sem layer)'}'"
    )

    records = []
    unrecognized_descriptions = set()
    for source_index, row in dataframe.iterrows():
        level_value = (
            row[level_column]
            if level_column is not None
            else default_nivel or "UNICO"
        )
        record = {
            "ITEM": str(row["ITEM"]).strip(),
            "NIVEL": normalizar_nivel(
                level_value,
                default=default_nivel or "UNICO",
            ),
            "CONCRETO_m3": 0.0,
            "FORMA_m2": 0.0,
        }

        for description_column, quantity_column in qto_pairs:
            description = normalize_desc(row[description_column])
            quantity = parse_qty(row[quantity_column])
            if not description or quantity == 0:
                continue

            normalized_description = _normalize_text(description)
            if "forma" in normalized_description:
                record["FORMA_m2"] += quantity
            elif "concreto" in normalized_description:
                record["CONCRETO_m3"] += quantity
            elif any(
                term in normalized_description
                for term in ("armadura", "aco", "ca50", "ca60")
            ):
                bitola = extract_bitola(description)
                if bitola is None:
                    raise ValueError(
                        "Armadura sem bitola reconhecida na linha "
                        f"{int(source_index) + 2}, ITEM '{record['ITEM']}': "
                        f"{description}"
                    )
                steel_key = f"ACO_{bitola}mm_kg"
                record[steel_key] = record.get(steel_key, 0.0) + quantity
            else:
                unrecognized_descriptions.add(description)

        records.append(record)

    result = pd.DataFrame(records)
    if result.empty:
        result = pd.DataFrame(
            columns=["ITEM", "NIVEL", "CONCRETO_m3", "FORMA_m2"]
        )

    steel_columns = sorted(
        [
            column
            for column in result.columns
            if _bitola_from_aco_column(column) is not None
        ],
        key=lambda column: _bitola_from_aco_column(column) or 0.0,
    )
    numeric_columns = ["CONCRETO_m3", "FORMA_m2"] + steel_columns
    for column in numeric_columns:
        if column not in result.columns:
            result[column] = 0.0
        result[column] = pd.to_numeric(
            result[column],
            errors="coerce",
        ).fillna(0.0)

    blank_item = result["ITEM"].map(_normalize_text).eq("")
    has_quantity = (
        result[numeric_columns]
        .abs()
        .sum(axis=1)
        .gt(0)
        if numeric_columns
        else pd.Series(False, index=result.index)
    )
    invalid_blank = blank_item & has_quantity
    if invalid_blank.any():
        raise ValueError(
            f"Foram encontradas {int(invalid_blank.sum())} linha(s) com "
            "quantitativos, mas com a coluna ITEM vazia."
        )
    result = result.loc[~blank_item].reset_index(drop=True)

    bitolas = [
        bitola
        for bitola in (
            _bitola_from_aco_column(column)
            for column in steel_columns
        )
        if bitola is not None
        and result[column].abs().sum() > 0
    ]
    ordered = [
        "ITEM",
        "NIVEL",
        "CONCRETO_m3",
        "FORMA_m2",
    ] + [
        column
        for column in steel_columns
        if _bitola_from_aco_column(column) in bitolas
    ]
    result = result[[column for column in ordered if column in result.columns]]

    if unrecognized_descriptions:
        examples = ", ".join(
            sorted(unrecognized_descriptions, key=_natural_sort_key)[:5]
        )
        print(
            "  AVISO: descrições não classificadas foram ignoradas: "
            f"{examples}"
        )

    return result, bitolas, has_level_column


def load_segmentos(segment_files):
    """
    Lê o mapa ITEM → SEGMENTO e bloqueia itens atribuídos a dois segmentos.
    """
    frames = []
    for segment_label, file_path in segment_files.items():
        dataframe = _read_excel_table(
            file_path,
            require_material_columns=False,
        )
        dataframe = _norm_cols(dataframe)
        dataframe = _drop_ignored_source_columns(dataframe)
        dataframe = dataframe.dropna(axis=1, how="all")

        columns = [str(column).strip() for column in dataframe.columns]
        item_column = _find_item_column(columns)
        if item_column is None:
            raise ValueError(
                f"Nenhuma coluna com 'item' nos 8 primeiros caracteres do título foi encontrada em "
                f"'{Path(file_path).name}'."
            )

        frame = pd.DataFrame(
            {
                "ITEM": _clean_text_series(dataframe[item_column]),
                "SEGMENTO": str(segment_label).strip(),
            }
        )
        frame = frame[frame["ITEM"].map(_normalize_text) != ""]
        frames.append(frame)

    if not frames:
        return pd.DataFrame(columns=["ITEM", "SEGMENTO"])

    combined = pd.concat(frames, ignore_index=True)
    conflicts = (
        combined.groupby("ITEM", dropna=False)["SEGMENTO"]
        .nunique()
        .loc[lambda series: series > 1]
    )
    if not conflicts.empty:
        examples = []
        for item in conflicts.index[:10]:
            segments = sorted(
                combined.loc[
                    combined["ITEM"] == item,
                    "SEGMENTO",
                ].unique(),
                key=_natural_sort_key,
            )
            examples.append(f"{item}: {', '.join(segments)}")
        raise ValueError(
            "Um mesmo ITEM foi encontrado em mais de um segmento:\n"
            + "\n".join(examples)
        )

    return combined.drop_duplicates("ITEM").reset_index(drop=True)


# ─── ESTILOS ─────────────────────────────────────────────────────────────────

def fill(c): return PatternFill("solid", start_color=c)

H_PILAR = fill("1F4E79")
H_CONC  = fill("1565C0")
H_FORMA = fill("1E6B3C")
H_ACO   = fill("4A235A")
H_TOTAL = fill("1F4E79")
H_SEM   = fill("C0392B")
S_FILLS = [fill("2E4057"), fill("1B4332"), fill("4A235A"),
           fill("7B241C"), fill("154360"), fill("0B5345")]
ALT     = [fill("F0F4F8"), fill("FFFFFF")]
L_CONC  = fill("D6E4F0")
L_FORMA = fill("D5E8D4")
L_ACO   = fill("EDE7F6")

WF = Font(name="Arial", bold=True, color="FFFFFF", size=10)
NF = Font(name="Arial", size=10)
BF = Font(name="Arial", bold=True, size=10)
C  = Alignment(horizontal="center", vertical="center", wrap_text=True)
L  = Alignment(horizontal="left",   vertical="center")
R  = Alignment(horizontal="right",  vertical="center")
thin = Side(style="thin", color="BFBFBF")
BRD  = Border(left=thin, right=thin, top=thin, bottom=thin)

def _c(ws, row, column, value, fnt=None, fl=None, al=None, fmt=None):
    if value is pd.NA or (
        isinstance(value, float)
        and pd.isna(value)
    ):
        value = None

    cell = ws.cell(row=row, column=column, value=value)
    cell.font = fnt or NF
    cell.fill = fl or PatternFill()
    cell.alignment = al or R
    cell.border = BRD
    if fmt:
        cell.number_format = fmt
    return cell

def _h(ws, r, c, v, fl=None):
    _c(ws, r, c, v, fnt=WF, fl=fl or H_PILAR, al=C)


# ─── WRITE HELPERS ───────────────────────────────────────────────────────────

def write_header(ws, extra_columns, bitolas):
    """Escreve o cabeçalho de duas linhas da tabela detalhada/resumida."""
    column_index = 1
    for label, _, header_fill in extra_columns:
        _h(ws, 1, column_index, label, fl=header_fill)
        _h(ws, 2, column_index, "", fl=header_fill)
        column_index += 1

    _h(ws, 1, column_index, "CONCRETO", fl=H_CONC)
    _h(ws, 2, column_index, "m3", fl=H_CONC)
    column_index += 1

    _h(ws, 1, column_index, "FORMA", fl=H_FORMA)
    _h(ws, 2, column_index, "m2", fl=H_FORMA)
    column_index += 1

    for bitola in bitolas:
        _h(ws, 1, column_index, f"O{bitola}mm", fl=H_ACO)
        _h(ws, 2, column_index, "kg", fl=H_ACO)
        column_index += 1

    ws.freeze_panes = "A3"
    return 3

def write_data_row(ws, row_index, row, extra_values, bitolas, row_fill):
    column_index = 1
    alignments = [L, C, C, C, C, C]

    for index, value in enumerate(extra_values):
        _c(
            ws,
            row_index,
            column_index,
            value,
            fnt=NF,
            fl=row_fill,
            al=(
                alignments[index]
                if index < len(alignments)
                else L
            ),
        )
        column_index += 1

    _c(
        ws,
        row_index,
        column_index,
        parse_qty(row.get("CONCRETO_m3", 0.0)),
        fnt=NF,
        fl=row_fill,
        al=R,
        fmt="#,##0.000",
    )
    column_index += 1
    _c(
        ws,
        row_index,
        column_index,
        parse_qty(row.get("FORMA_m2", 0.0)),
        fnt=NF,
        fl=row_fill,
        al=R,
        fmt="#,##0.000",
    )
    column_index += 1

    for bitola in bitolas:
        _c(
            ws,
            row_index,
            column_index,
            parse_qty(row.get(f"ACO_{bitola}mm_kg", 0.0)),
            fnt=NF,
            fl=row_fill,
            al=R,
            fmt="#,##0.000",
        )
        column_index += 1

def write_total_row(
    ws,
    row_index,
    labels,
    dataframe,
    bitolas,
    row_fill,
    extra_column_count,
):
    label_values = (
        list(labels)
        if isinstance(labels, (list, tuple))
        else [labels]
    )

    column_index = 1
    for position in range(extra_column_count):
        value = (
            label_values[position]
            if position < len(label_values)
            else ""
        )
        _c(
            ws,
            row_index,
            column_index,
            value,
            fnt=WF,
            fl=row_fill,
            al=L,
        )
        column_index += 1

    _c(
        ws,
        row_index,
        column_index,
        _series_sum(dataframe, "CONCRETO_m3"),
        fnt=WF,
        fl=row_fill,
        al=R,
        fmt="#,##0.000",
    )
    column_index += 1
    _c(
        ws,
        row_index,
        column_index,
        _series_sum(dataframe, "FORMA_m2"),
        fnt=WF,
        fl=row_fill,
        al=R,
        fmt="#,##0.000",
    )
    column_index += 1

    for bitola in bitolas:
        _c(
            ws,
            row_index,
            column_index,
            _series_sum(
                dataframe,
                f"ACO_{bitola}mm_kg",
            ),
            fnt=WF,
            fl=row_fill,
            al=R,
            fmt="#,##0.000",
        )
        column_index += 1


# ─── ABAS DINÂMICAS ──────────────────────────────────────────────────────────

def build_aba(
    workbook,
    sheet_title,
    header_title,
    dataframe,
    bitolas,
    summary_mode,
    total_fill,
    groups=None,
    label_column=None,
    summary_item="",
):
    """
    Cria uma aba completa ou resumida.

    No modo completo, ITEM contém o valor da coluna cujo título contém ``item``.
    No modo resumo, ITEM contém o nome da peça informado no front-end.
    """
    sheet = workbook.create_sheet(
        _safe_sheet_title(workbook, sheet_title)
    )
    data = dataframe.copy()

    for required in ("CONCRETO_m3", "FORMA_m2"):
        if required not in data.columns:
            data[required] = 0.0
    for bitola in bitolas:
        column = f"ACO_{bitola}mm_kg"
        if column not in data.columns:
            data[column] = 0.0

    valid_groups = {}
    if groups:
        valid_groups = {
            label: group.copy()
            for label, group in groups.items()
            if group is not None and not group.empty
        }

    if valid_groups and not summary_mode:
        extra_columns = [("ITEM", 20, H_PILAR)]
        if label_column and label_column != "ITEM":
            extra_columns.append((label_column, 14, H_PILAR))
            if (
                "NIVEL" in data.columns
                and label_column != "NIVEL"
            ):
                extra_columns.append(("NIVEL", 14, H_PILAR))
        elif "NIVEL" in data.columns:
            extra_columns.append(("NIVEL", 14, H_PILAR))

        row_index = write_header(sheet, extra_columns, bitolas)
        for group_index, (group_label, group_data) in enumerate(
            valid_groups.items()
        ):
            for item_index, (_, row) in enumerate(
                group_data.iterrows()
            ):
                extra_values = [row.get("ITEM", "")]
                if label_column and label_column != "ITEM":
                    extra_values.append(row.get(label_column, ""))
                if (
                    "NIVEL" in data.columns
                    and label_column != "NIVEL"
                ):
                    extra_values.append(row.get("NIVEL", ""))
                write_data_row(
                    sheet,
                    row_index,
                    row,
                    extra_values,
                    bitolas,
                    ALT[item_index % 2],
                )
                row_index += 1

            write_total_row(
                sheet,
                row_index,
                [f"SUBTOTAL {group_label}"],
                group_data,
                bitolas,
                S_FILLS[group_index % len(S_FILLS)],
                len(extra_columns),
            )
            row_index += 1

        write_total_row(
            sheet,
            row_index,
            ["TOTAL GERAL"],
            data,
            bitolas,
            H_TOTAL,
            len(extra_columns),
        )

    elif not summary_mode:
        extra_columns = [("ITEM", 20, H_PILAR)]
        if "NIVEL" in data.columns:
            extra_columns.append(("NIVEL", 14, H_PILAR))

        row_index = write_header(sheet, extra_columns, bitolas)
        for item_index, (_, row) in enumerate(data.iterrows()):
            extra_values = [row.get("ITEM", "")]
            if "NIVEL" in data.columns:
                extra_values.append(row.get("NIVEL", ""))
            write_data_row(
                sheet,
                row_index,
                row,
                extra_values,
                bitolas,
                ALT[item_index % 2],
            )
            row_index += 1

        write_total_row(
            sheet,
            row_index,
            ["TOTAL"],
            data,
            bitolas,
            total_fill,
            len(extra_columns),
        )

    else:
        piece_name = str(
            summary_item
            or header_title
            or "PEÇA"
        ).strip().upper()

        if valid_groups:
            group_header = label_column or "GRUPO"
            extra_columns = [
                ("ITEM", 30, H_PILAR),
                (group_header, 18, H_PILAR),
            ]
            row_index = write_header(
                sheet,
                extra_columns,
                bitolas,
            )
            for group_index, (group_label, group_data) in enumerate(
                valid_groups.items()
            ):
                write_total_row(
                    sheet,
                    row_index,
                    [piece_name, group_label],
                    group_data,
                    bitolas,
                    S_FILLS[group_index % len(S_FILLS)],
                    len(extra_columns),
                )
                row_index += 1

            write_total_row(
                sheet,
                row_index,
                [piece_name, "TOTAL GERAL"],
                data,
                bitolas,
                H_TOTAL,
                len(extra_columns),
            )
        else:
            extra_columns = [("ITEM", 30, H_PILAR)]
            row_index = write_header(
                sheet,
                extra_columns,
                bitolas,
            )
            write_total_row(
                sheet,
                row_index,
                [piece_name],
                data,
                bitolas,
                total_fill,
                len(extra_columns),
            )

    widths = [column[1] for column in extra_columns]
    widths += [13, 13] + [12] * len(bitolas)
    for column_index, width in enumerate(widths, 1):
        sheet.column_dimensions[
            get_column_letter(column_index)
        ].width = width

    print(f"  Aba '{sheet.title}': {len(data)} itens")
    return sheet

def build_total_geral(
    workbook,
    dataframe,
    bitolas,
    meta,
    group_labels=None,
    group_column=None,
):
    """Cria a tabela final com ITEM como primeira coluna."""
    requested_title = (
        f"TOTAL - {meta['elemento']} - {meta['segmento']}"
    )
    sheet = workbook.create_sheet(
        _safe_sheet_title(workbook, requested_title)
    )

    labels = list(group_labels) if group_labels else ["TOTAL"]
    headers = [
        "ITEM",
        "MATERIAL",
        "DESCRICAO",
        "TOTAL GERAL",
    ] + labels + ["UN"]

    observation = (
        f" - {meta['observacao']}"
        if meta.get("observacao")
        else ""
    )
    title_prefix = (
        "RESUMO DE MATERIAIS"
        if meta["modo_resumo"]
        else "LISTA DE MATERIAIS TOTAL"
    )
    title = (
        f"{title_prefix} - {meta['elemento']} - "
        f"{meta['pavimento']} - {meta['obra']} - "
        f"{meta['tipo']} - {meta['segmento']} - "
        f"{meta['data']}{observation}"
    )
    sheet.merge_cells(
        start_row=1,
        start_column=1,
        end_row=1,
        end_column=len(headers),
    )
    sheet["A1"].value = title
    sheet["A1"].font = Font(
        name="Arial",
        bold=True,
        size=12,
        color="1F4E79",
    )
    sheet["A1"].alignment = C
    sheet.row_dimensions[1].height = 22

    header_fills = [H_PILAR, H_PILAR, H_PILAR, H_TOTAL]
    header_fills += [
        S_FILLS[index % len(S_FILLS)]
        for index in range(len(labels))
    ]
    header_fills.append(H_PILAR)
    for column_index, (header, header_fill) in enumerate(
        zip(headers, header_fills),
        1,
    ):
        _h(sheet, 3, column_index, header, fl=header_fill)

    rows = [
        (
            "CONCRETO",
            "Concreto Estrutural",
            "CONCRETO_m3",
            "m3",
        ),
        (
            "FORMA",
            "Forma (Forma)",
            "FORMA_m2",
            "m2",
        ),
    ]
    for bitola in bitolas:
        rows.append(
            (
                "ACO",
                f"Armadura CA O {bitola} mm",
                f"ACO_{bitola}mm_kg",
                "kg",
            )
        )

    row_fills = {
        "CONCRETO": L_CONC,
        "FORMA": L_FORMA,
        "ACO": L_ACO,
    }
    piece_name = str(meta["elemento"]).strip().upper()
    row_index = 4

    for category, description, column, unit in rows:
        total = _series_sum(dataframe, column)
        if (
            group_labels
            and group_column
            and group_column in dataframe.columns
        ):
            group_values = [
                _series_sum(
                    dataframe[
                        dataframe[group_column] == group_label
                    ],
                    column,
                )
                for group_label in labels
            ]
        else:
            group_values = [total] * len(labels)

        values = [
            piece_name,
            category,
            description,
            total,
        ] + group_values + [unit]
        for column_index, value in enumerate(values, 1):
            numeric = 4 <= column_index <= 4 + len(labels)
            _c(
                sheet,
                row_index,
                column_index,
                value,
                fl=row_fills[category],
                al=(
                    L
                    if column_index in {1, 2, 3}
                    else R if numeric else C
                ),
                fmt="#,##0.000" if numeric else None,
            )
        row_index += 1

    steel_columns = [
        f"ACO_{bitola}mm_kg"
        for bitola in bitolas
    ]
    _c(
        sheet,
        row_index,
        1,
        piece_name,
        fnt=WF,
        fl=H_TOTAL,
        al=L,
    )
    _c(
        sheet,
        row_index,
        2,
        "TOTAL ACO",
        fnt=WF,
        fl=H_TOTAL,
        al=L,
    )
    _c(
        sheet,
        row_index,
        3,
        "",
        fnt=WF,
        fl=H_TOTAL,
        al=L,
    )
    _c(
        sheet,
        row_index,
        4,
        sum(
            _series_sum(dataframe, column)
            for column in steel_columns
        ),
        fnt=WF,
        fl=H_TOTAL,
        al=R,
        fmt="#,##0.000",
    )

    for column_index, group_label in enumerate(labels, start=5):
        if (
            group_labels
            and group_column
            and group_column in dataframe.columns
        ):
            group_data = dataframe[
                dataframe[group_column] == group_label
            ]
        else:
            group_data = dataframe
        group_total = sum(
            _series_sum(group_data, column)
            for column in steel_columns
        )
        _c(
            sheet,
            row_index,
            column_index,
            group_total,
            fnt=WF,
            fl=H_TOTAL,
            al=R,
            fmt="#,##0.000",
        )

    _c(
        sheet,
        row_index,
        5 + len(labels),
        "kg",
        fnt=WF,
        fl=H_TOTAL,
        al=C,
    )

    widths = [20, 12, 38, 16] + [14] * len(labels) + [8]
    for column_index, width in enumerate(widths, 1):
        sheet.column_dimensions[
            get_column_letter(column_index)
        ].width = width
    sheet.freeze_panes = "A4"
    print("  Aba 'TOTAL GERAL': resumo gerado")
    return sheet

def format_aba_nome(base, meta):
    if base == "TODOS":
        prefix = "RES- " if meta['modo_resumo'] else "TODOS - "
        return f"{prefix}{meta['elemento']} - {meta['segmento']}"[:31]
    if meta['modo_resumo'] and base.startswith("SEG"):
        return f"RES- {meta['elemento']} - {base}"[:31]
    return base[:31]

# ─── MAIN ────────────────────────────────────────────────────────────────────

def main():
    meta = coletar_metadados()
    file_qto, segment_files, file_out = coletar_arquivos(meta)

    print("Carregando QTO...")
    data, bitolas, has_level_column = load_qto(
        file_qto,
        default_nivel=meta["pavimento"],
    )
    if data.empty:
        raise ValueError(
            "A planilha não possui itens válidos para gerar a lista."
        )
    print(f"  {len(data)} itens | bitolas: {bitolas}")

    if has_level_column:
        file_levels = sorted(
            data["NIVEL"].dropna().astype(str).unique(),
            key=_natural_sort_key,
        )
        print(f"  Níveis no arquivo: {file_levels}")
        same_single_level = (
            len(file_levels) == 1
            and _normalize_text(file_levels[0])
            == _normalize_text(meta["pavimento"])
        )
        if not same_single_level:
            preview = ", ".join(file_levels[:4])
            if len(file_levels) > 4:
                preview += "..."
            choice = _ask_choice(
                "Fonte do pavimento",
                f"Arquivo tem coluna Layer/Nível com "
                f"{len(file_levels)} nível(is).\n"
                "Qual usar como pavimento?",
                [
                    f"Usar Layer/Nível do arquivo ({preview})",
                    (
                        "Usar pavimento informado: "
                        f"{meta['pavimento']} (1 nível único)"
                    ),
                ],
            )
            if "informado" in choice:
                data["NIVEL"] = meta["pavimento"]
                print(
                    f"  Nível forçado para '{meta['pavimento']}'"
                )
            else:
                print("  Usando níveis da planilha")

    if segment_files:
        print("Carregando segmentos...")
        segment_map = load_segmentos(segment_files)
        data = data.merge(
            segment_map,
            on="ITEM",
            how="left",
            validate="many_to_one",
        )
        missing_segment_count = int(data["SEGMENTO"].isna().sum())
        if missing_segment_count:
            print(
                f"  AVISO: {missing_segment_count} itens sem segmento "
                "-> 'SEM SEGMENTO'"
            )
        data["SEGMENTO"] = data["SEGMENTO"].fillna(
            "SEM SEGMENTO"
        )
    else:
        data["SEGMENTO"] = "TODOS"

    data["NIVEL"] = data["NIVEL"].apply(
        lambda value: normalizar_nivel(
            value,
            default=meta["pavimento"],
        )
    )
    data["SEGMENTO"] = _clean_text_series(
        data["SEGMENTO"]
    ).replace("", "SEM SEGMENTO")

    all_segments = sorted(
        data["SEGMENTO"].unique(),
        key=_natural_sort_key,
    )
    regular_segments = [
        segment
        for segment in all_segments
        if segment != "SEM SEGMENTO"
    ]
    levels = sorted(
        data["NIVEL"].unique(),
        key=_natural_sort_key,
    )
    has_unassigned_segment = "SEM SEGMENTO" in all_segments
    multiple_levels = len(levels) > 1

    print(f"  Segmentos : {all_segments}")
    print(f"  Níveis    : {levels}")

    workbook = Workbook()
    if workbook.active is not None:
        workbook.remove(workbook.active)

    if segment_files:
        for segment_index, segment in enumerate(regular_segments):
            segment_data = data[
                data["SEGMENTO"] == segment
            ].reset_index(drop=True)
            if segment_data.empty:
                continue

            segment_levels = sorted(
                segment_data["NIVEL"].unique(),
                key=_natural_sort_key,
            )
            if len(segment_levels) > 1:
                groups = {}
                for level in segment_levels:
                    level_data = segment_data[
                        segment_data["NIVEL"] == level
                    ].reset_index(drop=True)
                    if level_data.empty:
                        continue
                    groups[level] = level_data
                    build_aba(
                        workbook,
                        f"{segment} - {level}",
                        "ITEM",
                        level_data,
                        bitolas,
                        meta["modo_resumo"],
                        S_FILLS[segment_index % len(S_FILLS)],
                        summary_item=meta["elemento"],
                    )

                build_aba(
                    workbook,
                    format_aba_nome(segment, meta),
                    "ITEM",
                    segment_data,
                    bitolas,
                    meta["modo_resumo"],
                    S_FILLS[segment_index % len(S_FILLS)],
                    groups=groups,
                    label_column="NIVEL",
                    summary_item=meta["elemento"],
                )
            else:
                build_aba(
                    workbook,
                    format_aba_nome(segment, meta),
                    "ITEM",
                    segment_data,
                    bitolas,
                    meta["modo_resumo"],
                    S_FILLS[segment_index % len(S_FILLS)],
                    summary_item=meta["elemento"],
                )
    elif multiple_levels:
        for level in levels:
            level_data = data[
                data["NIVEL"] == level
            ].reset_index(drop=True)
            if level_data.empty:
                continue
            build_aba(
                workbook,
                str(level),
                "ITEM",
                level_data,
                bitolas,
                meta["modo_resumo"],
                S_FILLS[0],
                summary_item=meta["elemento"],
            )
    else:
        build_aba(
            workbook,
            format_aba_nome("TODOS", meta),
            "ITEM",
            data,
            bitolas,
            meta["modo_resumo"],
            H_TOTAL,
            summary_item=meta["elemento"],
        )

    if has_unassigned_segment:
        unassigned = data[
            data["SEGMENTO"] == "SEM SEGMENTO"
        ].reset_index(drop=True)
        build_aba(
            workbook,
            "SEM SEGMENTO",
            "ITEM",
            unassigned,
            bitolas,
            meta["modo_resumo"],
            H_SEM,
            summary_item=meta["elemento"],
        )
        print(
            f"  ATENÇÃO: {len(unassigned)} itens em "
            "'SEM SEGMENTO'"
        )

    if segment_files and len(all_segments) > 1:
        group_labels = all_segments
        group_column = "SEGMENTO"
    elif not segment_files and multiple_levels:
        group_labels = levels
        group_column = "NIVEL"
    else:
        group_labels = None
        group_column = None

    build_total_geral(
        workbook,
        data,
        bitolas,
        meta,
        group_labels,
        group_column,
    )

    _save_workbook(workbook, file_out)
    _copy_output_to_cache(file_out)

    if _ask_yes_no(
        "Planilha gerada",
        f"Planilha gerada:\n{file_out.name}\n\n"
        "Deseja abrir a planilha agora?",
    ):
        _open_output_file(file_out)

    print(f"\nSalvo em:\n{file_out}")
    print(f"\nTotal : {len(data)} itens")
    print(f"Bitolas: {bitolas}")
    for segment in all_segments:
        count = int((data["SEGMENTO"] == segment).sum())
        print(f"  {segment}: {count} itens")
    for level in levels:
        count = int((data["NIVEL"] == level).sum())
        print(f"  {level}: {count} itens")

    return _ask_yes_no(
        "Concluído",
        f"Planilha gerada:\n{file_out.name}\n\n"
        "Deseja gerar uma NOVA planilha?\n\n"
        "SIM = novo processo     NÃO = finalizar",
    )


if __name__ == "__main__":
    ENV_VARS = [
        "LM_ELEMENTO",
        "LM_PAVIMENTO",
        "LM_SEGMENTO",
        "LM_OBRA",
        "LM_TIPO",
        "LM_DATA",
        "LM_MODO",
        "LM_OBS",
    ]

    while True:
        try:
            action = _ask_main_action()
            if action == 1:
                continue_processing = main()
                if not continue_processing:
                    print("\nFinalizado. Até logo!")
                    break
            elif action == 2:
                merge_material_lists()
                if not _ask_yes_no(
                    "Menu",
                    "Deseja voltar ao menu principal?",
                ):
                    print("\nFinalizado. Até logo!")
                    break
            else:
                print("\nCancelado. Até logo!")
                break
        except SystemExit as exc:
            print(str(exc) or "Operação cancelada pelo usuário.")
        except Exception as exc:
            _show_unexpected_error(exc)
        finally:
            for variable in ENV_VARS:
                os.environ.pop(variable, None)
