# Gestão de Stock de Cápsulas

Aplicação web para gerir o inventário de cápsulas de uma forma simples e intuitiva.

## 🚀 Início Rápido

### Opção 1: Se tem Python instalado

Clique duplo em **`run.bat`** — é tudo!

### Opção 2: Se não tem Python

1. Execute **`install.bat`** (instala Python automaticamente)
2. Depois execute **`run.bat`**

A app abrirá automaticamente em `http://localhost:8501`

## 📋 Como Usar

- **🔄 Recarregar Excel** — sincroniza com o ficheiro .xlsx na pasta
- **➕ Nova Cápsula** — adiciona novo registo manualmente
- **📤 Exportar Excel** — descarrega todos os dados para Excel
- **✏️ Editar/Remover** — selecione o ID do registo
- **🔍 Pesquisa e Filtros** — encontre registos rapidamente

## 📁 Ficheiros

| Ficheiro           | Descrição                              |
| ------------------ | -------------------------------------- |
| `run.bat`          | **Clique aqui para correr a app**      |
| `install.bat`      | Instala Python (se necessário)         |
| `app.py`           | Código da aplicação                    |
| `requirements.txt` | Dependências Python                    |
| `README.md`        | Este ficheiro                          |
| `capsulas.db`      | Base de dados (criada automaticamente) |
| `*.xlsx`           | Ficheiros Excel para importar          |

## ❓ Resolução de Problemas

### "Python não encontrado"

- Execute `install.bat` para instalar Python automaticamente
- Ou instale manualmente: https://www.python.org/downloads/
- **Importante:** selecione "Add Python to PATH" durante a instalação

### "Porta 8501 já em uso"

- Feche outras instâncias da aplicação
- Ou altere a porta em `run.bat` (mudar `8501` para outro número, ex: `8502`)

### "Erro ao instalar dependências"

- Desconecte da VPN (pode bloquear pip)
- Execute `run.bat` novamente
- Ou execute manualmente: `pip install -r requirements.txt`

## 💾 Dados

- Os dados são guardados automaticamente em `capsulas.db`
- Registos adicionados manualmente ficam preservados ao recarregar Excel
- Sempre que recarregar, o Excel é sincronizado com a BD

## 📝 Notas

- A aplicação está otimizada para Firefox e Chrome
- Funciona em qualquer Windows 10+
- Sem necessidade de instalação — apenas execute `run.bat`
