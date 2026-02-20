# Preenche Planilhas

Aplicativo que lê um **PDF de Plano de Aplicação** e gera uma **planilha Excel preenchida de Itens** automaticamente.
<img width="624" height="860" alt="image" src="https://github.com/user-attachments/assets/b6d29437-27dd-4762-8923-61b279c62caf" />


## Como funciona
1. Você envia o PDF do plano.
2. O app extrai os itens e suas informações (meta, item, artigo, bem/serviço, instituição, natureza, quantidade, unidade, valor total).
3. É gerada uma planilha com os campos preenchidos e pronta para download.
4. Caso existam campos vazios, o app lista **as células em branco**.

## Estrutura do projeto
- `app.py` — interface Streamlit
- `preencher_planilha.py` — lógica de extração e geração do Excel
- `Planilha Base.xlsx` — template da planilha
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
- O template `Planilha Base.xlsx` deve estar na mesma pasta do app.
- O PDF deve seguir o padrão de “META ESPECÍFICA” e “Item” para extração correta.

## Colaboração
- Este repositório pode ser público para consulta, download e fork.
- Contribuições de terceiros devem ser feitas por fork + Pull Request.

## Licença
Este projeto está licenciado sob a licença MIT. Veja `LICENSE`.
