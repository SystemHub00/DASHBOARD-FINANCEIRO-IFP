# IFP Dashboard Financeiro

Dashboard financeiro gerencial para o Instituto de Formação Profissional.

## Funcionalidades

- Upload drag-and-drop de planilhas .xlsx (contas pagas e recebidas)
- KPIs: Receita Total, Despesa Total, Resultado Líquido, Ticket Médio
- Fluxo de caixa diário (barras comparativas)
- Saldo acumulado (linha)
- Receitas por categoria (doughnut)
- Top 15 despesas por categoria (barras horizontais)
- Distribuição por banco (receitas e despesas)
- Meios de pagamento
- Ranking top 10 fornecedores/beneficiários
- Tabela interativa com filtros por categoria, banco e busca livre

## Como rodar localmente

```bash
pip install -r requirements.txt
python app.py
```

Acesse: http://localhost:5000

## Deploy no Render

1. Faça push do código para um repositório GitHub/GitLab
2. No Render, clique em **New → Web Service**
3. Conecte o repositório
4. Defina:
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `gunicorn app:app --bind 0.0.0.0:$PORT --workers 1 --timeout 120`
5. Adicione a variável de ambiente:
   - `FLASK_SECRET_KEY` → qualquer string aleatória longa
6. Clique em **Deploy**

## Formato das Planilhas

As planilhas devem ter as colunas:
`Id | Vencimento | Competência | Previsto para | Data de pagamento | CPF/CNPJ | Nome | Descrição | Referência | Categoria | Detalhamento | Centro de Custo | Valor categoria/centro de custo | Identificador | Banco`

A planilha de receitas pode ter ainda a coluna `Número NFS-e`.
