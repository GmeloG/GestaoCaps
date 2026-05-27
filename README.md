# Gestão de Stock de Cápsulas

Aplicação Windows para gerir o inventário de cápsulas de forma simples e intuitiva.

---

## Arquitetura

| O quê | Onde fica |
|---|---|
| `GestaoCapsulas.exe` | PC local de cada utilizador (Desktop ou pasta local) |
| `capsulas.db` | Pasta de rede partilhada |
| Ficheiros `.xlsx` | Pasta de rede partilhada |

**Pasta de rede:**
```
\\sidel.com\emea\pt-smf\groups\STORAGE\Máquinas\Produção\caps
```

---

## Instalação para o utilizador

1. Copiar `GestaoCapsulas.exe` para o Desktop (ou outra pasta local)
2. Fazer duplo-clique para abrir

A app verifica automaticamente a ligação à rede antes de abrir. Se a rede não estiver disponível, mostra um aviso e não abre.

---

## Como Usar a App

| Ação | Como fazer |
|---|---|
| **Reload Excel** | Importa o ficheiro `.xlsx` da pasta de rede |
| **Nova Cápsula** | Abre formulário para adicionar manualmente |
| **Exportar Excel** | Guarda os dados num ficheiro Excel |
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
4. Distribuir `dist\GestaoCapsulas.exe` pelos utilizadores (Desktop local)
5. Garantir que `capsulas.db` existe na pasta de rede

Para correr a app localmente sem gerar o `.exe` (requer ligação à rede):

```
run.bat
```

---

## Ficheiros do projeto

| Ficheiro | Descrição |
|---|---|
| `app.py` | Código da aplicação |
| `build.bat` | Gera o `.exe` |
| `run.bat` | Corre a app localmente (requer Python) |
| `install.bat` | Instala Python e dependências |
| `requirements.txt` | Dependências Python |

---

## Dados e Segurança

- `capsulas.db` fica sempre na pasta de rede — nunca no PC local
- Registos adicionados manualmente ficam preservados ao recarregar o Excel
- Fazer backup regular do ficheiro `capsulas.db`
- Apenas um utilizador deve usar a app de cada vez

---

## Resolução de Problemas

**"Rede não disponível" ao abrir**
- Verificar se está ligado à rede da empresa (VPN ou rede local)
- Confirmar que a pasta de rede está acessível no Explorador de Ficheiros

**Antivírus bloqueia o ficheiro**
- Adicionar exceção para `GestaoCapsulas.exe` no antivírus da empresa

**Erro ao gerar o .exe**
- Verificar que o Python está instalado: correr `install.bat`
- Desligar VPN durante o build se houver erros de rede
