# Gestão de Stock de Cápsulas

Aplicação para gerir o inventário de cápsulas de forma simples e intuitiva.

---

## Utilização (pasta partilhada)

**Os colegas só precisam de fazer duplo-clique em `GestaoCapsulas.exe`.**

- Uma janela preta abre — é normal, **não fechar**
- O browser abre automaticamente em `http://localhost:8501`
- A base de dados (`capsulas.db`) fica sempre na mesma pasta que o `.exe`

> Se o antivírus bloquear o ficheiro, é necessário adicionar uma exceção para `GestaoCapsulas.exe`.

---

## Como Usar a App

| Botão | Função |
|---|---|
| **Reload** | Sincroniza com o ficheiro `.xlsx` na pasta |
| **Nova Capsula** | Adiciona um registo manualmente |
| **Exportar Excel** | Descarrega todos os dados para Excel |
| **Editar / Remover** | Seleciona o ID do registo na tabela |
| **Pesquisa e Filtros** | Encontra registos rapidamente |

---

## Para o Administrador (gerar o .exe)

Necessário apenas uma vez, ou quando houver alterações ao código.

**Requisitos:** Python instalado com PyInstaller (`pip install pyinstaller`)

1. Abrir a pasta do projeto no PC com Python
2. Fazer duplo-clique em **`build.bat`**
3. Aguardar 3-5 minutos
4. Copiar para a pasta partilhada:
   - `dist\GestaoCapsulas.exe`
   - `capsulas.db` (se já tiver dados)

Para correr a app localmente sem gerar o `.exe`:

```
run.bat
```

---

## Ficheiros

| Ficheiro | Descrição |
|---|---|
| `GestaoCapsulas.exe` | **Executável para a pasta partilhada** |
| `capsulas.db` | Base de dados (criada automaticamente) |
| `build.bat` | Gera o `.exe` a partir do código |
| `run.bat` | Corre a app localmente (requer Python) |
| `launcher.py` | Ponto de entrada do PyInstaller |
| `app.py` | Código da aplicação |
| `requirements.txt` | Dependências Python |
| `*.xlsx` | Ficheiros Excel para importar |

---

## Dados e Segurança

- Os dados são guardados em `capsulas.db`, na mesma pasta que o `.exe`
- Registos adicionados manualmente ficam preservados ao recarregar o Excel
- Fazer backup regular do ficheiro `capsulas.db`
- Apenas um utilizador deve usar a app de cada vez

---

## Resolução de Problemas

**O browser não abre automaticamente**
- Abrir manualmente: `http://localhost:8501`

**"Porta 8501 já em uso"**
- Fechar outras instâncias da aplicação (verificar se a janela preta ainda está aberta)

**Antivírus bloqueia o ficheiro**
- Adicionar exceção para `GestaoCapsulas.exe` no antivírus da empresa

**Erro ao gerar o .exe**
- Verificar que o PyInstaller está instalado: `pip install pyinstaller`
- Desligar VPN durante o build
