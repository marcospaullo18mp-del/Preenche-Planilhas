# Preenche Planilhas

Aplicativo Streamlit que lê um **PDF de Plano de Aplicação** e gera uma **planilha Excel preenchida** automaticamente.

## Como funciona
1. Você envia o PDF do plano.
2. O app extrai os itens e suas informações (meta, item, artigo, bem/serviço, instituição, natureza, quantidade, unidade, valor total).
3. É gerada uma planilha com os campos preenchidos e pronta para download.
4. Caso existam campos vazios, o app lista **as células em branco**.

## Estrutura do projeto
- `app.py` — interface Streamlit
- `preencher_planilha.py` — lógica de extração e geração do Excel
- `Itens NT.xlsx` — template da planilha
- `Logo.png` — logo exibida no topo
- `requirements.txt` — dependências

## Rodar localmente
```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

## Como usar
1. Abra o app.
2. Faça upload do PDF do Plano de Aplicação.
3. Clique em **Processar**.
4. Baixe a planilha gerada no botão **Baixar planilha**.

## Observações
- O template `Itens NT.xlsx` deve estar na mesma pasta do app.
- O PDF deve seguir o padrão de “META ESPECÍFICA” e “Item” para extração correta.

## Deploy gratuito (Streamlit Cloud)
1. Suba este projeto no GitHub.
2. No Streamlit Community Cloud, crie um novo app apontando para o repositório.
3. Selecione `app.py` como arquivo principal.

Pronto! O app ficará disponível online.
