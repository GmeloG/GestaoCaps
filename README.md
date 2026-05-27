# Gestão de Stock de Cápsulas

Aplicação Windows para gerir o inventário de cápsulas de forma simples e intuitiva.

---

## Utilização (pasta partilhada)

**Os colegas só precisam de fazer duplo-clique em `GestaoCapsulas.exe`.**

- Abre diretamente como uma janela Windows normal — sem browser, sem instalação
- A base de dados (`capsulas.db`) fica sempre na mesma pasta que o `.exe`

> Se o antivírus bloquear o ficheiro, é necessário adicionar uma exceção para `GestaoCapsulas.exe`.

---

## Como Usar a App

| Ação | Como fazer |
|---|---|
| **Reload Excel** | Importa o ficheiro `.xlsx` da mesma pasta |
| **Nova Cápsula** | Abre formulário para adicionar manualmente |
| **Exportar Excel** | Guarda todos os dados num ficheiro Excel |
| **Editar registo** | Duplo-clique na linha, ou selecionar ID e clicar Editar |
| **Apagar registo** | Selecionar ID e clicar Apagar (pede confirmação) |
| **Pesquisa** | Escrever na caixa de pesquisa (filtra em tempo real) |
| **Filtrar por Estado** | Selecionar estado no filtro |
| **Ordenar** | Clicar no cabeçalho de qualquer coluna |
| **Menu de contexto** | Botão direito do rato numa linha |

---

## Para o Administrador (gerar o .exe)

Necessário apenas uma vez, ou quando houver alterações ao código.

**Requisitos:** Python com as dependências instaladas (`install.bat`)

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
| `install.bat` | Instala Python e dependências |
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

**Antivírus bloqueia o ficheiro**
- Adicionar exceção para `GestaoCapsulas.exe` no antivírus da empresa

**Erro ao gerar o .exe**
- Verificar que o Python está instalado: correr `install.bat`
- O build usa uma pasta temporária local para evitar conflitos com o OneDrive
- Desligar VPN durante o build se houver erros de rede
