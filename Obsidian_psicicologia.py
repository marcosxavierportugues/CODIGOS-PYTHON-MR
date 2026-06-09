from pathlib import Path
import pyperclip

# ALTERE PARA O CAMINHO DO SEU VAULT
vault = Path(r"D:\OSIDIAN BRAIN - D")

pasta = vault / "8- Pessoal"
pasta.mkdir(parents=True, exist_ok=True)

arquivo = pasta / "pscicologia.md"

texto = pyperclip.paste()

with open(arquivo, "a", encoding="utf-8") as f:
    f.write("\n\n---\n\n")
    f.write(texto)

print("Nota salva:", arquivo)