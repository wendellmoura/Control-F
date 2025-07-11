# Control+F Busca Avançada em Arquivos Excel/CSV

**Control+F** é uma aplicação desktop em Python para busca avançada de termos em arquivos Excel (`.xlsx`, `.xls`) e CSV, com interface gráfica moderna baseada em [ttkbootstrap](https://github.com/israel-dryer/ttkbootstrap). Permite localizar valores, selecionar colunas para exportação e gerar relatórios em diversos formatos.

## Funcionalidades

- **Busca instantânea** em arquivos Excel ou CSV, em uma aba ou em todas as abas (worksheets).
- **Visualização dos resultados** com navegação rápida.
- **Seleção avançada de colunas** para exportar apenas os dados relevantes.
- **Exportação dos resultados** para JSON, CSV e Excel.
- **Log integrado** das ações e erros.
- Suporte a arquivos grandes e diferentes delimitadores em CSV.

## Captura de tela
<img width="787" height="607" alt="image" src="https://github.com/user-attachments/assets/38511148-8994-4c49-8480-0bd80f543ba0" />

## Instalação

1. **Clone o repositório** ou baixe o arquivo principal.

2. **Instale as dependências** (recomenda-se uso de ambiente virtual):

```bash
pip install pandas ttkbootstrap
```

3. Execute o programa:

```bash
python seu_arquivo.py
```

## Uso

- **Selecione um arquivo** (Excel ou CSV) usando o botão "Selecionar".
- **Clique em "Carregar"** para detectar as abas (worksheets).
- **Escolha a aba** para busca ou utilize "Buscar em todas as abas".
- **Digite o termo de busca** e clique em "Buscar".
- Os resultados aparecerão na tabela.
- Use "Exportar colunas" para selecionar apenas as colunas desejadas para exportação.
- Exporte os resultados para JSON, CSV ou Excel conforme necessário.
- Use os botões de limpeza para reiniciar a busca ou apagar o log.

## Principais Classes

- `LocalFileSearcher`: Carrega arquivos, detecta abas e realiza buscas nos dados.
- `ColumnSelector`: Janela modal para seleção personalizada de colunas.
- `FileSearchApp`: Gerencia a interface gráfica principal e integra as funcionalidades.

## Requisitos

- Python 3.8+
- [pandas](https://pandas.pydata.org/)
- [ttkbootstrap](https://github.com/israel-dryer/ttkbootstrap)
- Suporte nativo a arquivos `.csv`, `.xlsx`, `.xls`.

## Observações

- O programa detecta automaticamente o delimitador de arquivos CSV.
- Suporte a exportação customizada por colunas, inclusive em arquivos Excel.
- Os dados nunca são enviados para a nuvem — tudo é processado localmente.

## Licença

Este projeto é distribuído sob a licença MIT.

## Autor

Wendell Moura

---

*Pull requests e sugestões são bem-vindos!*
