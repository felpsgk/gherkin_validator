# gherkin_validator
Validador de Cenários Gherkin (.feature)

# 🧪 Validador de Cenários Gherkin (.feature)

Este projeto tem como objetivo padronizar e validar cenários escritos no formato Gherkin (`.feature`) utilizados em automações de testes. Desenvolvido em Python, o validador detecta problemas estruturais, ordenação incorreta de palavras-chave, estrutura de exemplos e muito mais — além de gerar relatórios `.docx` e `.xlsx` com os erros encontrados.

---

## 🚀 Funcionalidades

- 🔍 Leitura recursiva de arquivos `.feature`
- 📚 Parser com a biblioteca oficial `gherkin-official`
- ✅ Validação de estrutura, ordem e uso das keywords (Given, When, Then, And, etc.)
- 📊 Geração de relatórios em `.docx` e `.xlsx`
- 📈 Criação de relatório de cobertura de automação com base em tags (`@automatizado`, `@automatizar`, `@manual`)

---

## 🧰 Tecnologias Utilizadas

- [gherkin-official](https://pypi.org/project/gherkin-official/)
- [python-docx](https://python-docx.readthedocs.io/)
- [openpyxl](https://openpyxl.readthedocs.io/)
- [pandas](https://pandas.pydata.org/)
- [argparse](https://docs.python.org/3/library/argparse.html)

---

## 💻 Instalação

1. Clone o repositório:
```bash
git clone https://github.com/seu-usuario/validador-gherkin.git
cd validador-gherkin
