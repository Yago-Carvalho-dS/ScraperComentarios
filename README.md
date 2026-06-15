# ScraperComentarios

Automação para extração e consolidação de comentários de pesquisas (Qualtrics/Survey) utilizando Selenium e Pandas.

## Descrição

Este projeto automatiza o processo de login, navegação e extração de dados de comentários de arquivos Excel baixados de plataformas de pesquisa. Ele filtra especificamente colunas que contenham a palavra "comente" e consolida tudo em um único arquivo CSV.

## Requisitos

- Python 3.x
- Selenium
- Pandas
- Webdriver Manager
- Openpyxl (para leitura de arquivos Excel)

## Instalação

1. Clone o repositório:
   ```bash
   git clone https://github.com/Yago-Carvalho-dS/ScraperComentarios.git
   ```
2. Instale as dependências:
   ```bash
   pip install -r requirements.txt
   ```

## Uso

Execute o script principal:
```bash
python Auto.py
```

## Estrutura do Projeto

- `Auto.py`: Script principal de automação e processamento.
- `Labels.txt`: Possivelmente contém rótulos ou configurações auxiliares.
- `imgs/`: Capturas de tela para referência.
