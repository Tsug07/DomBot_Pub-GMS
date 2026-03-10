# DomBot_Pub_GMS

![Python](https://img.shields.io/badge/Python-3.8%2B-3776AB?logo=python&logoColor=white)
![Platform](https://img.shields.io/badge/Platform-Windows-0078D6?logo=windows&logoColor=white)
![CustomTkinter](https://img.shields.io/badge/UI-CustomTkinter-blue)
![License](https://img.shields.io/badge/License-Propriet%C3%A1rio-red)
![Status](https://img.shields.io/badge/Status-Em%20Produ%C3%A7%C3%A3o-brightgreen)

Automação desktop (RPA) para publicação em lote de guias GMS no sistema **Domínio Folha**, com interface gráfica moderna em modo escuro.

---

## Funcionalidades

- Publicação automatizada de documentos externos no Domínio Folha
- Leitura de planilha Excel com os dados de cada documento (Nº, Período, Salvar Como, Caminho)
- Validação do arquivo Excel antes da execução
- Barra de progresso em tempo real
- Log detalhado de cada etapa da execução
- Detecção automática da janela do Domínio Folha
- Botão de parada com interrupção segura a qualquer momento
- Registro de logs em arquivo (`publicacao_log.txt`)

## Pré-requisitos

- **Windows 10** ou superior
- **Python 3.8+**
- **Domínio Folha** instalado e aberto com a tela de *Publicação de Documentos Externos* visível

## Instalação

```bash
git clone https://github.com/seu-usuario/DomBot_Pub-GMS.git
cd DomBot_Pub-GMS
pip install -r requirements.txt
```

### Dependências

| Pacote | Finalidade |
|---|---|
| `customtkinter` | Interface gráfica moderna |
| `pandas` | Leitura de planilhas Excel |
| `openpyxl` | Engine para leitura de `.xlsx` |
| `pywinauto` | Automação de interface Windows |
| `pywin32` | Interação com a API Win32 |
| `Pillow` | Carregamento de imagens/logo |

## Uso

1. Abra o **Domínio Folha** e navegue até a tela de **Publicação de Documentos Externos**.
2. Execute o DomBot:
   ```bash
   python DomBot_Pub_GMS.py
   ```
   Ou utilize o atalho `iniciar_DomBot_Pub.bat`.
3. Clique em **Selecionar Excel** e escolha a planilha com os dados.
4. (Opcional) Clique em **Validar Excel** para verificar a integridade do arquivo.
5. Clique em **Publicar** para iniciar o processo automatizado.

## Formato da Planilha Excel

A planilha deve conter as seguintes colunas obrigatórias:

| Coluna | Descrição |
|---|---|
| `Nº` | Número do documento |
| `Periodo` | Período de referência |
| `Salvar Como` | Nome de identificação do documento |
| `Caminho` | Caminho completo do arquivo PDF a ser publicado |

## Estrutura do Projeto

```
DomBot_Pub-GMS/
├── assets/
│   ├── DomBot_Pub.ico        # Ícone da aplicação
│   └── DomBot_Pub.png        # Logo exibida na interface
├── logs/                     # Logs de execução (ignorado pelo git)
│   └── publicacao_YYYYMMDD_HHMMSS.log
├── DomBot_Pub_GMS.py         # Código principal
├── iniciar_DomBot_Pub.bat    # Script de inicialização rápida
├── .gitignore
└── README.md
```

## Screenshot

A interface possui tema escuro com logo, seletor de arquivos, botões de ação, barra de progresso e painel de log integrado.

---

## Autor

**Hugo L. Almeida**

---

> Desenvolvido para automatizar processos repetitivos de publicação de documentos no sistema Domínio Folha, aumentando a produtividade e reduzindo erros manuais.
