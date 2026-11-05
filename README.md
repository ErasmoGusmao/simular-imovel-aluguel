# Simulador de valor de imóvel e aluguel

## 🔜 Migração futura: ambiente reprodutível com `uv`

> Planejada, **ainda não feita** neste projeto (registrada em 05/11/2026).

**Situação atual:** não há arquivo de dependências; o código roda no Python/Anaconda da máquina. Bibliotecas externas encontradas nos imports: `numpy`, `pandas` (conferir).

**Por que migrar:**

- Um `requirements.txt` com faixas (`pandas>=2.0`), ou o Python/Anaconda compartilhado da máquina,
  guarda só a **intenção**: cada instalação pode trazer versões diferentes — principalmente das
  dependências transitivas (numpy, pytz, tzdata…). É o clássico "na minha máquina funciona".
- O `uv` separa os dois papéis em arquivos distintos, ambos versionados:
  - `pyproject.toml` — **o que o projeto precisa**: dependências diretas com faixa de versão,
    editado por você (`uv add` / `uv remove`);
  - `uv.lock` — **o que foi resolvido**: versão exata e hash de **todas** as dependências, para
    Windows e Linux, gerado pelo `uv` (nunca editado à mão);
  - `.python-version` — a versão do Python do projeto (o `uv` a baixa se faltar).
- `uv sync` deixa o `.venv` **idêntico** ao lock: instala o que falta e **remove o que sobra**
  (o `pip install -r` só acrescenta). O `.venv` vira descartável — recriar leva segundos.
- Dependências só de desenvolvimento (`pytest`, `jupyter`, `ipykernel`…) ficam separadas das de
  execução (`[dependency-groups]`).
- O `requirements.txt` não some: passa a ser **gerado** do lock (`uv export`) para quem só tem `pip`.

**Como migrar (quando for a hora):**

```bash
winget install --id=astral-sh.uv -e      # uma vez por máquina (Linux: curl -LsSf https://astral.sh/uv/install.sh | sh)
uv init --bare                            # cria só o pyproject.toml
uv python pin 3.13                        # .python-version
uv add numpy pandas   # conferir a lista e pôr faixas de versão
uv lock && uv sync                        # gera o uv.lock e cria o .venv
uv export --format requirements-txt --no-dev --no-hashes -o requirements.txt
```

Versionar `pyproject.toml`, `uv.lock` e `.python-version` (e o `requirements.txt` gerado); o
`.venv` fica fora do Git. Projeto no `D:` com o cache do `uv` no `C:`: `setx UV_LINK_MODE copy`.
Documentação: <https://docs.astral.sh/uv/>.
