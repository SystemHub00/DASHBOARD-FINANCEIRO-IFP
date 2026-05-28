import os
import json
import uuid
from io import BytesIO
from datetime import datetime
from collections import defaultdict

import openpyxl
from flask import (
    Flask, render_template_string, request, redirect, url_for,
    session, jsonify, flash
)

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "ifp-dashboard-2026-secret")

# Registrar a função abs como global no Jinja2
app.jinja_env.globals.update(abs=abs)

UPLOAD_FOLDER = os.path.join(os.path.dirname(__file__), "uploads")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# ─────────────────────────────────────────────────────────────
# HELPERS DE LEITURA
# ─────────────────────────────────────────────────────────────
def read_excel(path):
    wb = openpyxl.load_workbook(path, read_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    if not rows:
        return [], []
    headers = [str(h) if h is not None else "" for h in rows[0]]
    data = []
    for row in rows[1:]:
        obj = {}
        for i, h in enumerate(headers):
            val = row[i] if i < len(row) else None
            obj[h] = val
        data.append(obj)
    return headers, data


def fmt_currency(val):
    return f"R$ {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def parse_data(despesas_rows, receitas_rows):
    """Calcula todos os KPIs e séries para o dashboard."""

    # ── Totais gerais ──────────────────────────────────────────
    total_receita  = sum(abs(float(r["Valor categoria/centro de custo"] or 0)) for r in receitas_rows if r.get("Valor categoria/centro de custo"))
    total_despesa  = sum(abs(float(r["Valor categoria/centro de custo"] or 0)) for r in despesas_rows if r.get("Valor categoria/centro de custo"))
    resultado      = total_receita - total_despesa
    margem         = (resultado / total_receita * 100) if total_receita else 0

    # ── Receitas por categoria ──────────────────────────────────
    rec_por_cat = defaultdict(float)
    for r in receitas_rows:
        cat = r.get("Categoria") or r.get("Nome") or "Outros"
        val = abs(float(r.get("Valor categoria/centro de custo") or 0))
        rec_por_cat[cat] += val

    # ── Despesas por categoria (top 15) ────────────────────────
    desp_por_cat = defaultdict(float)
    for r in despesas_rows:
        cat = r.get("Categoria") or "Outros"
        val = abs(float(r.get("Valor categoria/centro de custo") or 0))
        desp_por_cat[cat] += val
    desp_por_cat_sorted = sorted(desp_por_cat.items(), key=lambda x: -x[1])[:15]

    # ── Por banco ──────────────────────────────────────────────
    rec_por_banco = defaultdict(float)
    for r in receitas_rows:
        banco = (r.get("Banco") or "Não informado").strip()
        val   = abs(float(r.get("Valor categoria/centro de custo") or 0))
        rec_por_banco[banco] += val

    desp_por_banco = defaultdict(float)
    for r in despesas_rows:
        banco = (r.get("Banco") or "Não informado").strip()
        val   = abs(float(r.get("Valor categoria/centro de custo") or 0))
        desp_por_banco[banco] += val

    # ── Top fornecedores / beneficiários ───────────────────────
    top_fornecedores = defaultdict(float)
    for r in despesas_rows:
        nome = r.get("Nome") or "Não informado"
        val  = abs(float(r.get("Valor categoria/centro de custo") or 0))
        top_fornecedores[nome] += val
    top_fornecedores_list = sorted(top_fornecedores.items(), key=lambda x: -x[1])[:10]

    # ── Fluxo de caixa diário ──────────────────────────────────
    daily_rec  = defaultdict(float)
    daily_desp = defaultdict(float)
    for r in receitas_rows:
        dt  = str(r.get("Data de pagamento") or r.get("Vencimento") or "")
        val = abs(float(r.get("Valor categoria/centro de custo") or 0))
        if dt:
            daily_rec[dt] += val
    for r in despesas_rows:
        dt  = str(r.get("Data de pagamento") or r.get("Vencimento") or "")
        val = abs(float(r.get("Valor categoria/centro de custo") or 0))
        if dt:
            daily_desp[dt] += val

    all_dates = sorted(set(list(daily_rec.keys()) + list(daily_desp.keys())))
    daily_labels   = all_dates
    daily_rec_vals  = [round(daily_rec.get(d, 0), 2) for d in all_dates]
    daily_desp_vals = [round(daily_desp.get(d, 0), 2) for d in all_dates]

    # Saldo acumulado
    saldo_acum = []
    acc = 0
    for r, d in zip(daily_rec_vals, daily_desp_vals):
        acc += r - d
        saldo_acum.append(round(acc, 2))

    # ── Receitas por forma de pagamento (banco como proxy) ─────
    metodos = {
        "Itaú": 0, "Celcoin - Conta Corrente": 0,
        "Cartão Débito/Crédito": 0, "Caixa Espécie": 0,
        "Banco do Brasil": 0, "Outros": 0
    }
    for r in receitas_rows:
        banco = (r.get("Banco") or "Outros").strip()
        val   = abs(float(r.get("Valor categoria/centro de custo") or 0))
        matched = False
        for k in metodos:
            if k.lower() in banco.lower():
                metodos[k] += val
                matched = True
                break
        if not matched:
            metodos["Outros"] += val
    metodos = {k: v for k, v in metodos.items() if v > 0}

    # ── Quantidade de transações ────────────────────────────────
    qtd_receitas  = len(receitas_rows)
    qtd_despesas  = len(despesas_rows)

    # ── Ticket médio ────────────────────────────────────────────
    mensalidades = [abs(float(r.get("Valor categoria/centro de custo") or 0))
                    for r in receitas_rows
                    if "mensalidade" in str(r.get("Categoria") or "").lower()]
    ticket_medio = sum(mensalidades) / len(mensalidades) if mensalidades else 0

    # ── Matrículas ──────────────────────────────────────────────
    matriculas_val = sum(abs(float(r.get("Valor categoria/centro de custo") or 0))
                         for r in receitas_rows
                         if "matrícula" in str(r.get("Categoria") or "").lower()
                         or "matricula" in str(r.get("Categoria") or "").lower())

    return {
        # KPIs
        "total_receita":       total_receita,
        "total_despesa":       total_despesa,
        "resultado":           resultado,
        "margem":              margem,
        "qtd_receitas":        qtd_receitas,
        "qtd_despesas":        qtd_despesas,
        "ticket_medio":        ticket_medio,
        "matriculas_val":      matriculas_val,
        # Gráficos
        "rec_por_cat":         dict(rec_por_cat),
        "desp_por_cat":        dict(desp_por_cat_sorted),
        "rec_por_banco":       dict(rec_por_banco),
        "desp_por_banco":      dict(desp_por_banco),
        "metodos_pagamento":   metodos,
        "top_fornecedores":    top_fornecedores_list,
        "daily_labels":        daily_labels,
        "daily_rec_vals":      daily_rec_vals,
        "daily_desp_vals":     daily_desp_vals,
        "saldo_acum":          saldo_acum,
        # Tabelas
        "despesas_rows":       despesas_rows,
        "receitas_rows":       receitas_rows,
    }


# ─────────────────────────────────────────────────────────────
# TEMPLATES
# ─────────────────────────────────────────────────────────────
UPLOAD_TEMPLATE = r"""
<!DOCTYPE html>
<html lang="pt-br">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>IFP Dashboard Financeiro</title>
<style>
  :root {
    --blue: #003087; --blue-light: #0047AB; --blue-soft: #EEF3FF;
    --red: #CC0000;  --red-light: #FF1A1A; --red-soft: #FFF0F0;
    --white: #FFFFFF; --gray: #F5F7FA; --gray2: #E8ECF0;
    --text: #1A1A2E;  --text2: #4A5568;
    --shadow: 0 4px 24px rgba(0,48,135,0.12);
    --radius: 16px;
  }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Segoe UI', system-ui, sans-serif; background: var(--gray); color: var(--text); min-height: 100vh; }

  /* HEADER */
  .header {
    background: var(--blue);
    padding: 0 32px;
    display: flex; align-items: center; justify-content: space-between;
    height: 68px; box-shadow: 0 2px 12px rgba(0,0,0,0.2);
  }
  .header-brand { display: flex; align-items: center; gap: 16px; }
  .header-logo {
    width: 52px; height: 52px; border-radius: 10px;
    background: var(--white); display: flex; align-items: center; justify-content: center;
    font-size: 1.4rem; font-weight: 900; color: var(--blue); letter-spacing: -1px;
  }
  .header-title { color: var(--white); font-size: 1.25rem; font-weight: 700; }
  .header-subtitle { color: rgba(255,255,255,0.7); font-size: 0.82rem; margin-top: 2px; }
  .header-badge {
    background: var(--red); color: var(--white);
    padding: 6px 14px; border-radius: 999px;
    font-size: 0.78rem; font-weight: 700; letter-spacing: 0.04em;
    text-transform: uppercase;
  }

  /* MAIN */
  .main { max-width: 860px; margin: 0 auto; padding: 48px 16px; }
  .page-title {
    text-align: center; margin-bottom: 8px;
    font-size: 2rem; font-weight: 800; color: var(--blue);
  }
  .page-sub { text-align: center; color: var(--text2); margin-bottom: 40px; font-size: 1rem; }

  /* FLASH */
  .flash { background: #fff3cd; border: 1px solid #ffc107; border-radius: 10px; padding: 12px 18px; margin-bottom: 24px; color: #856404; font-size: 0.95rem; }
  .flash.error { background: #f8d7da; border-color: #f5c6cb; color: #721c24; }

  /* UPLOAD CARD */
  .upload-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 24px; margin-bottom: 32px; }
  @media(max-width:600px){ .upload-grid { grid-template-columns: 1fr; } }
  .upload-card {
    background: var(--white); border-radius: var(--radius);
    border: 2px dashed var(--gray2); padding: 32px 24px;
    text-align: center; transition: all 0.2s;
    cursor: pointer; position: relative;
  }
  .upload-card:hover, .upload-card.dragover {
    border-color: var(--blue-light);
    background: var(--blue-soft);
    transform: translateY(-2px);
    box-shadow: var(--shadow);
  }
  .upload-card.receitas:hover, .upload-card.receitas.dragover { border-color: var(--blue); }
  .upload-card.despesas:hover, .upload-card.despesas.dragover { border-color: var(--red); }
  .upload-card input[type=file] {
    position: absolute; inset: 0; opacity: 0; cursor: pointer; width: 100%; height: 100%;
  }
  .upload-icon { font-size: 3rem; margin-bottom: 14px; }
  .upload-card h3 { font-size: 1.1rem; font-weight: 700; margin-bottom: 8px; }
  .upload-card.receitas h3 { color: var(--blue); }
  .upload-card.despesas h3 { color: var(--red); }
  .upload-card p { font-size: 0.88rem; color: var(--text2); line-height: 1.5; }
  .upload-card .file-name {
    margin-top: 12px; padding: 8px 14px; border-radius: 8px;
    background: var(--gray); font-size: 0.82rem; color: var(--text2);
    display: none; word-break: break-all;
  }
  .upload-card .file-name.visible { display: block; }
  .file-ok { background: #d4edda !important; color: #155724 !important; font-weight: 700; }

  /* SUBMIT */
  .submit-area { text-align: center; }
  .btn-submit {
    background: linear-gradient(135deg, var(--blue) 0%, var(--blue-light) 100%);
    color: var(--white); border: none; border-radius: 12px;
    padding: 16px 48px; font-size: 1.1rem; font-weight: 700;
    cursor: pointer; box-shadow: 0 6px 20px rgba(0,48,135,0.3);
    transition: all 0.2s; letter-spacing: 0.02em; text-transform: uppercase;
  }
  .btn-submit:hover { transform: translateY(-2px); box-shadow: 0 10px 28px rgba(0,48,135,0.4); }
  .btn-submit:active { transform: translateY(0); }

  /* INFO */
  .info-section { margin-top: 48px; }
  .info-grid { display: grid; grid-template-columns: repeat(3,1fr); gap: 16px; margin-top: 24px; }
  @media(max-width:600px){ .info-grid { grid-template-columns: 1fr; } }
  .info-card {
    background: var(--white); border-radius: 12px; padding: 20px;
    border-left: 4px solid var(--blue); box-shadow: 0 2px 12px rgba(0,0,0,0.06);
  }
  .info-card.red { border-left-color: var(--red); }
  .info-card h4 { font-size: 0.9rem; font-weight: 700; color: var(--blue); margin-bottom: 6px; }
  .info-card.red h4 { color: var(--red); }
  .info-card p { font-size: 0.83rem; color: var(--text2); line-height: 1.5; }

  .footer { text-align: center; padding: 32px; color: var(--text2); font-size: 0.82rem; }
  .footer span { color: var(--blue); font-weight: 700; }
</style>
</head>
<body>
<header class="header">
  <div class="header-brand">
    <div class="header-logo">IFP</div>
    <div>
      <div class="header-title">Instituto de Formação Profissional</div>
      <div class="header-subtitle">Dashboard Financeiro Gerencial</div>
    </div>
  </div>
  <div class="header-badge">⬆ Upload de Arquivos</div>
</header>

<main class="main">
  <h1 class="page-title">📊 Painel Financeiro IFP</h1>
  <p class="page-sub">Importe suas planilhas de contas pagas e recebidas para visualizar o resultado da empresa em tempo real.</p>

  {% for msg, cat in messages %}
  <div class="flash {{ 'error' if cat == 'error' else '' }}">{{ msg }}</div>
  {% endfor %}

  <form method="POST" action="/upload" enctype="multipart/form-data" id="upload-form">
    <div class="upload-grid">
      <div class="upload-card receitas" id="card-receitas">
        <input type="file" name="receitas" id="file-receitas" accept=".xlsx" required>
        <div class="upload-icon">💰</div>
        <h3>Contas Recebidas</h3>
        <p>Arraste ou clique para importar o arquivo de <strong>receitas</strong> (.xlsx)</p>
        <div class="file-name" id="name-receitas">Nenhum arquivo selecionado</div>
      </div>
      <div class="upload-card despesas" id="card-despesas">
        <input type="file" name="despesas" id="file-despesas" accept=".xlsx" required>
        <div class="upload-icon">📋</div>
        <h3>Contas Pagas</h3>
        <p>Arraste ou clique para importar o arquivo de <strong>despesas</strong> (.xlsx)</p>
        <div class="file-name" id="name-despesas">Nenhum arquivo selecionado</div>
      </div>
    </div>
    <div class="submit-area">
      <button type="submit" class="btn-submit">🚀 Gerar Dashboard</button>
    </div>
  </form>

  <div class="info-section">
    <div class="info-grid">
      <div class="info-card">
        <h4>📑 Formato Esperado</h4>
        <p>Arquivo .xlsx com colunas: Id, Vencimento, Competência, Data de pagamento, CPF/CNPJ, Nome, Categoria, Valor, Banco.</p>
      </div>
      <div class="info-card">
        <h4>📊 O que você verá</h4>
        <p>KPIs financeiros, fluxo de caixa, análise por categoria, ranking de fornecedores e muito mais.</p>
      </div>
      <div class="info-card red">
        <h4>🔒 Segurança</h4>
        <p>Os dados são processados localmente no servidor e não são armazenados permanentemente após a sessão.</p>
      </div>
    </div>
  </div>
</main>

<footer class="footer">
  <span>IFP</span> — Instituto de Formação Profissional · Dashboard Financeiro v1.0
</footer>

<script>
function setupUpload(cardId, inputId, nameId) {
  const card  = document.getElementById(cardId);
  const input = document.getElementById(inputId);
  const nameEl = document.getElementById(nameId);

  function showName(file) {
    nameEl.textContent = '✅ ' + file.name;
    nameEl.classList.add('visible', 'file-ok');
  }

  input.addEventListener('change', function() {
    if (this.files[0]) showName(this.files[0]);
  });
  card.addEventListener('dragover', e => { e.preventDefault(); card.classList.add('dragover'); });
  card.addEventListener('dragleave', () => card.classList.remove('dragover'));
  card.addEventListener('drop', e => {
    e.preventDefault(); card.classList.remove('dragover');
    const file = e.dataTransfer.files[0];
    if (file && file.name.endsWith('.xlsx')) {
      const dt = new DataTransfer();
      dt.items.add(file);
      input.files = dt.files;
      showName(file);
    }
  });
}
setupUpload('card-receitas','file-receitas','name-receitas');
setupUpload('card-despesas','file-despesas','name-despesas');
</script>
</body>
</html>
"""

DASHBOARD_TEMPLATE = r"""
<!DOCTYPE html>
<html lang="pt-br">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>IFP Dashboard Financeiro</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
  :root {
    --blue: #003087; --blue-light: #0047AB; --blue-mid: #1565C0;
    --blue-soft: #EEF3FF; --blue-pale: #F5F8FF;
    --red: #CC0000;  --red-soft: #FFF0F0; --red-mid: #E53935;
    --white: #FFFFFF; --gray: #F4F6FA; --gray2: #E2E8F0; --gray3: #CBD5E1;
    --text: #1A202C;  --text2: #4A5568; --text3: #718096;
    --green: #0A7C3E; --green-soft: #E8F5E9;
    --shadow-sm: 0 2px 8px rgba(0,48,135,0.08);
    --shadow: 0 4px 20px rgba(0,48,135,0.12);
    --shadow-lg: 0 8px 40px rgba(0,48,135,0.16);
    --radius: 16px; --radius-sm: 10px;
    --transition: all 0.25s ease;
  }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  html { scroll-behavior: smooth; }
  body { font-family: 'Segoe UI', system-ui, -apple-system, sans-serif; background: var(--gray); color: var(--text); line-height: 1.5; }

  /* ── HEADER ── */
  .header {
    background: var(--blue); position: sticky; top: 0; z-index: 100;
    padding: 0 24px; height: 64px;
    display: flex; align-items: center; justify-content: space-between;
    box-shadow: 0 2px 16px rgba(0,0,0,0.25);
  }
  .header-brand { display: flex; align-items: center; gap: 14px; }
  .header-logo {
    width: 44px; height: 44px; border-radius: 10px; background: var(--white);
    display: flex; align-items: center; justify-content: center;
    font-size: 1.1rem; font-weight: 900; color: var(--blue); line-height: 1;
    flex-shrink: 0;
  }
  .header-title { color: var(--white); font-size: 1.1rem; font-weight: 700; }
  .header-subtitle { color: rgba(255,255,255,0.65); font-size: 0.78rem; }
  .header-actions { display: flex; align-items: center; gap: 12px; }
  .btn-new-upload {
    background: rgba(255,255,255,0.15); border: 1px solid rgba(255,255,255,0.3);
    color: var(--white); padding: 8px 16px; border-radius: 999px;
    font-size: 0.82rem; font-weight: 600; cursor: pointer;
    text-decoration: none; transition: var(--transition);
    display: flex; align-items: center; gap: 6px;
  }
  .btn-new-upload:hover { background: rgba(255,255,255,0.25); }
  .header-date { color: rgba(255,255,255,0.6); font-size: 0.78rem; }

  /* ── LAYOUT ── */
  .container { max-width: 1400px; margin: 0 auto; padding: 24px 16px 48px; }

  /* ── SECTION TITLES ── */
  .section-header { display: flex; align-items: center; gap: 10px; margin: 32px 0 16px; }
  .section-header h2 { font-size: 1.15rem; font-weight: 700; color: var(--blue); }
  .section-divider { flex: 1; height: 2px; background: linear-gradient(90deg, var(--blue-soft), transparent); border-radius: 999px; }
  .section-badge {
    background: var(--blue-soft); color: var(--blue);
    padding: 4px 12px; border-radius: 999px; font-size: 0.75rem; font-weight: 700;
  }

  /* ── KPI CARDS ── */
  .kpi-grid { display: grid; grid-template-columns: repeat(4,1fr); gap: 16px; margin-bottom: 8px; }
  @media(max-width:1100px){ .kpi-grid { grid-template-columns: repeat(2,1fr); } }
  @media(max-width:560px)  { .kpi-grid { grid-template-columns: 1fr; } }
  .kpi-card {
    background: var(--white); border-radius: var(--radius); padding: 22px 20px;
    box-shadow: var(--shadow-sm); position: relative; overflow: hidden;
    border-top: 4px solid transparent; transition: var(--transition);
  }
  .kpi-card:hover { transform: translateY(-3px); box-shadow: var(--shadow); }
  .kpi-card.receita  { border-top-color: var(--blue); }
  .kpi-card.despesa  { border-top-color: var(--red); }
  .kpi-card.resultado { border-top-color: {% if data.resultado >= 0 %}#0A7C3E{% else %}#CC0000{% endif %}; }
  .kpi-card.extra    { border-top-color: #7B61FF; }
  .kpi-icon {
    width: 44px; height: 44px; border-radius: 12px;
    display: flex; align-items: center; justify-content: center;
    font-size: 1.4rem; margin-bottom: 14px; flex-shrink: 0;
  }
  .kpi-card.receita  .kpi-icon { background: var(--blue-soft); }
  .kpi-card.despesa  .kpi-icon { background: var(--red-soft); }
  .kpi-card.resultado .kpi-icon { background: {% if data.resultado >= 0 %}var(--green-soft){% else %}var(--red-soft){% endif %}; }
  .kpi-card.extra    .kpi-icon { background: #F3F0FF; }
  .kpi-label { font-size: 0.78rem; font-weight: 600; color: var(--text3); text-transform: uppercase; letter-spacing: 0.06em; margin-bottom: 6px; }
  .kpi-value { font-size: 1.55rem; font-weight: 800; color: var(--text); line-height: 1.1; }
  .kpi-card.receita  .kpi-value { color: var(--blue); }
  .kpi-card.despesa  .kpi-value { color: var(--red); }
  .kpi-card.resultado .kpi-value { color: {% if data.resultado >= 0 %}var(--green){% else %}var(--red){% endif %}; }
  .kpi-card.extra    .kpi-value { color: #7B61FF; }
  .kpi-sub { font-size: 0.8rem; color: var(--text3); margin-top: 4px; }

  /* Faixa decorativa */
  .kpi-card::after {
    content: ''; position: absolute; right: -18px; top: -18px;
    width: 80px; height: 80px; border-radius: 50%;
    opacity: 0.06;
  }
  .kpi-card.receita::after  { background: var(--blue); }
  .kpi-card.despesa::after  { background: var(--red); }
  .kpi-card.resultado::after { background: var(--green); }
  .kpi-card.extra::after    { background: #7B61FF; }

  /* ── CHART CARDS ── */
  .chart-grid-2 { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }
  .chart-grid-3 { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 20px; }
  .chart-grid-13 { display: grid; grid-template-columns: 1fr 2fr; gap: 20px; }
  .chart-grid-31 { display: grid; grid-template-columns: 2fr 1fr; gap: 20px; }
  @media(max-width:900px){
    .chart-grid-2, .chart-grid-3, .chart-grid-13, .chart-grid-31
    { grid-template-columns: 1fr; }
  }
  .chart-card {
    background: var(--white); border-radius: var(--radius); padding: 24px;
    box-shadow: var(--shadow-sm); transition: var(--transition);
  }
  .chart-card:hover { box-shadow: var(--shadow); }
  .chart-card.full-width { grid-column: 1 / -1; }
  .chart-title {
    font-size: 0.95rem; font-weight: 700; color: var(--text); margin-bottom: 4px;
    display: flex; align-items: center; gap: 8px;
  }
  .chart-sub { font-size: 0.78rem; color: var(--text3); margin-bottom: 18px; }
  .chart-container { position: relative; height: 280px; }
  .chart-container.tall { height: 360px; }
  .chart-container.short { height: 220px; }

  /* ── TABELAS ── */
  .table-card {
    background: var(--white); border-radius: var(--radius); padding: 24px;
    box-shadow: var(--shadow-sm); overflow: hidden;
  }
  .table-filters {
    display: flex; gap: 12px; flex-wrap: wrap; margin-bottom: 16px; align-items: center;
  }
  .filter-input {
    border: 1.5px solid var(--gray2); border-radius: 8px;
    padding: 8px 14px; font-size: 0.85rem; color: var(--text);
    outline: none; background: var(--gray); transition: var(--transition);
    min-width: 160px;
  }
  .filter-input:focus { border-color: var(--blue-light); background: var(--white); }
  .filter-label { font-size: 0.78rem; color: var(--text3); font-weight: 600; }
  .table-wrapper { overflow-x: auto; border-radius: var(--radius-sm); }
  table { width: 100%; border-collapse: collapse; font-size: 0.83rem; }
  thead th {
    background: var(--blue); color: var(--white);
    padding: 11px 14px; text-align: left; font-weight: 600;
    font-size: 0.78rem; text-transform: uppercase; letter-spacing: 0.04em;
    white-space: nowrap;
  }
  thead th:first-child { border-radius: 8px 0 0 0; }
  thead th:last-child  { border-radius: 0 8px 0 0; }
  tbody tr { border-bottom: 1px solid var(--gray2); transition: background 0.15s; }
  tbody tr:hover { background: var(--blue-pale); }
  tbody td { padding: 10px 14px; color: var(--text2); white-space: nowrap; }
  tbody td.valor-pos { color: var(--green); font-weight: 700; }
  tbody td.valor-neg { color: var(--red);   font-weight: 700; }
  .badge {
    display: inline-block; padding: 3px 10px; border-radius: 999px;
    font-size: 0.73rem; font-weight: 700; letter-spacing: 0.02em;
    white-space: nowrap;
  }
  .badge-blue { background: var(--blue-soft); color: var(--blue); }
  .badge-red  { background: var(--red-soft);  color: var(--red); }
  .badge-green{ background: var(--green-soft); color: var(--green); }
  .badge-gray { background: var(--gray2);     color: var(--text3); }

  /* ── FORNECEDORES ── */
  .fornec-list { display: flex; flex-direction: column; gap: 8px; }
  .fornec-item {
    display: flex; align-items: center; gap: 12px;
    padding: 10px 14px; border-radius: 10px; background: var(--gray);
    transition: var(--transition);
  }
  .fornec-item:hover { background: var(--blue-soft); }
  .fornec-rank {
    width: 26px; height: 26px; border-radius: 50%;
    background: var(--blue); color: var(--white);
    font-size: 0.75rem; font-weight: 800;
    display: flex; align-items: center; justify-content: center;
    flex-shrink: 0;
  }
  .fornec-rank.top3 { background: var(--red); }
  .fornec-name { flex: 1; font-size: 0.85rem; font-weight: 600; color: var(--text); min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  .fornec-val  { font-size: 0.85rem; font-weight: 800; color: var(--red); white-space: nowrap; }
  .fornec-bar  { width: 100%; height: 4px; background: var(--gray2); border-radius: 999px; margin-top: 3px; }
  .fornec-bar-fill { height: 100%; border-radius: 999px; background: linear-gradient(90deg, var(--blue), var(--blue-light)); }

  /* ── RESULTADO BANNER ── */
  .resultado-banner {
    border-radius: var(--radius); padding: 20px 28px;
    display: flex; align-items: center; justify-content: space-between;
    flex-wrap: wrap; gap: 16px; margin-bottom: 8px;
    {% if data.resultado >= 0 %}
    background: linear-gradient(135deg, #E8F5E9, #C8E6C9); border-left: 5px solid var(--green);
    {% else %}
    background: linear-gradient(135deg, var(--red-soft), #FFCDD2); border-left: 5px solid var(--red);
    {% endif %}
  }
  .resultado-banner .label { font-size: 0.85rem; font-weight: 600; color: var(--text2); }
  .resultado-banner .value { font-size: 1.8rem; font-weight: 900; {% if data.resultado >= 0 %}color: var(--green);{% else %}color: var(--red);{% endif %} }
  .resultado-banner .pct   { font-size: 0.9rem; color: var(--text2); }

  /* ── PAGINAÇÃO ── */
  .pagination { display: flex; gap: 8px; align-items: center; justify-content: flex-end; margin-top: 16px; flex-wrap: wrap; }
  .page-btn {
    padding: 6px 12px; border-radius: 8px; border: 1px solid var(--gray2);
    background: var(--white); color: var(--text2); font-size: 0.82rem;
    cursor: pointer; transition: var(--transition);
  }
  .page-btn.active, .page-btn:hover { background: var(--blue); color: var(--white); border-color: var(--blue); }
  .page-info { font-size: 0.78rem; color: var(--text3); }

  /* ── FOOTER ── */
  .footer { text-align: center; padding: 24px; color: var(--text3); font-size: 0.78rem; border-top: 1px solid var(--gray2); margin-top: 32px; }

  /* ── TOOLTIP CUSTOMIZADO ── */
  .tab-btns { display: flex; gap: 8px; margin-bottom: 16px; }
  .tab-btn {
    padding: 7px 16px; border-radius: 8px; border: 1.5px solid var(--gray2);
    background: var(--white); color: var(--text2); font-size: 0.82rem;
    font-weight: 600; cursor: pointer; transition: var(--transition);
  }
  .tab-btn.active { background: var(--blue); color: var(--white); border-color: var(--blue); }
  .tab-panel { display: none; }
  .tab-panel.active { display: block; }
</style>
</head>
<body>

<header class="header">
  <div class="header-brand">
    <div class="header-logo">IFP</div>
    <div>
      <div class="header-title">Instituto de Formação Profissional</div>
      <div class="header-subtitle">Dashboard Financeiro Gerencial</div>
    </div>
  </div>
  <div class="header-actions">
    <span class="header-date" id="header-date"></span>
    <a href="/" class="btn-new-upload">↑ Novo Upload</a>
  </div>
</header>

<div class="container">

  <!-- ── RESULTADO BANNER ── -->
  <div class="resultado-banner">
    <div>
      <div class="label">RESULTADO DO PERÍODO</div>
      <div class="value">{{ fmt(data.resultado) }}</div>
    </div>
    <div>
      <div class="label">Margem Líquida</div>
      <div class="value" style="font-size:1.4rem;">{{ "%.1f"|format(data.margem) }}%</div>
    </div>
    <div>
      <div class="label">Receita Total</div>
      <div class="pct">{{ fmt(data.total_receita) }}</div>
    </div>
    <div>
      <div class="label">Despesa Total</div>
      <div class="pct">{{ fmt(data.total_despesa) }}</div>
    </div>
    <div>
      <div class="label">Transações</div>
      <div class="pct">{{ data.qtd_receitas + data.qtd_despesas }} registros</div>
    </div>
  </div>

  <!-- ── KPIs ── -->
  <div class="section-header">
    <h2>📌 Indicadores Chave</h2>
    <div class="section-divider"></div>
    <span class="section-badge">KPIs</span>
  </div>
  <div class="kpi-grid">
    <div class="kpi-card receita">
      <div class="kpi-icon">💰</div>
      <div class="kpi-label">Receita Total</div>
      <div class="kpi-value">{{ fmt(data.total_receita) }}</div>
      <div class="kpi-sub">{{ data.qtd_receitas }} transações de entrada</div>
    </div>
    <div class="kpi-card despesa">
      <div class="kpi-icon">📋</div>
      <div class="kpi-label">Despesa Total</div>
      <div class="kpi-value">{{ fmt(data.total_despesa) }}</div>
      <div class="kpi-sub">{{ data.qtd_despesas }} transações de saída</div>
    </div>
    <div class="kpi-card resultado">
      <div class="kpi-icon">{% if data.resultado >= 0 %}📈{% else %}📉{% endif %}</div>
      <div class="kpi-label">Resultado Líquido</div>
      <div class="kpi-value">{{ fmt(data.resultado) }}</div>
      <div class="kpi-sub">Margem de {{ "%.1f"|format(data.margem) }}%</div>
    </div>
    <div class="kpi-card extra">
      <div class="kpi-icon">🎓</div>
      <div class="kpi-label">Ticket Médio Mensalidade</div>
      <div class="kpi-value">{{ fmt(data.ticket_medio) }}</div>
      <div class="kpi-sub">Receita matrículas: {{ fmt(data.matriculas_val) }}</div>
    </div>
  </div>

  <!-- ── FLUXO DE CAIXA ── -->
  <div class="section-header">
    <h2>📆 Fluxo de Caixa Diário</h2>
    <div class="section-divider"></div>
    <span class="section-badge">Evolução temporal</span>
  </div>
  <div class="chart-card">
    <div class="chart-title">💵 Receitas × Despesas por Dia</div>
    <div class="chart-sub">Comparativo diário do período importado</div>
    <div class="chart-container tall">
      <canvas id="chartFluxo"></canvas>
    </div>
  </div>

  <!-- ── SALDO ACUMULADO ── -->
  <div style="margin-top:20px;">
    <div class="chart-card">
      <div class="chart-title">📊 Saldo Acumulado (Caixa)</div>
      <div class="chart-sub">Evolução do caixa acumulado dia a dia</div>
      <div class="chart-container">
        <canvas id="chartSaldo"></canvas>
      </div>
    </div>
  </div>

  <!-- ── CATEGORIAS ── -->
  <div class="section-header">
    <h2>🏷️ Análise por Categoria</h2>
    <div class="section-divider"></div>
    <span class="section-badge">Composição</span>
  </div>
  <div class="chart-grid-2">
    <div class="chart-card">
      <div class="chart-title">💰 Receitas por Categoria</div>
      <div class="chart-sub">Distribuição percentual das entradas</div>
      <div class="chart-container">
        <canvas id="chartRecCat"></canvas>
      </div>
    </div>
    <div class="chart-card">
      <div class="chart-title">📋 Top Despesas por Categoria</div>
      <div class="chart-sub">Maiores categorias de custo</div>
      <div class="chart-container tall">
        <canvas id="chartDespCat"></canvas>
      </div>
    </div>
  </div>

  <!-- ── BANCOS & MEIOS ── -->
  <div class="section-header">
    <h2>🏦 Bancos & Meios de Pagamento</h2>
    <div class="section-divider"></div>
    <span class="section-badge">Distribuição</span>
  </div>
  <div class="chart-grid-3">
    <div class="chart-card">
      <div class="chart-title">💳 Receitas por Banco</div>
      <div class="chart-sub">Canal de recebimento</div>
      <div class="chart-container short">
        <canvas id="chartRecBanco"></canvas>
      </div>
    </div>
    <div class="chart-card">
      <div class="chart-title">🏦 Despesas por Banco</div>
      <div class="chart-sub">Canal de pagamento</div>
      <div class="chart-container short">
        <canvas id="chartDespBanco"></canvas>
      </div>
    </div>
    <div class="chart-card">
      <div class="chart-title">💵 Meios de Pagamento (Receitas)</div>
      <div class="chart-sub">Pix, Cartão, Espécie e outros</div>
      <div class="chart-container short">
        <canvas id="chartMetodos"></canvas>
      </div>
    </div>
  </div>

  <!-- ── TOP FORNECEDORES ── -->
  <div class="section-header">
    <h2>👥 Top Fornecedores / Beneficiários</h2>
    <div class="section-divider"></div>
    <span class="section-badge">Ranking</span>
  </div>
  <div class="chart-grid-13">
    <div class="chart-card">
      <div class="chart-title">🏆 Ranking de Gastos</div>
      <div class="chart-sub">Top 10 maiores recebedores de pagamentos</div>
      <div class="fornec-list" id="fornec-list"></div>
    </div>
    <div class="chart-card">
      <div class="chart-title">📊 Distribuição Visual</div>
      <div class="chart-sub">Participação no total de despesas</div>
      <div class="chart-container">
        <canvas id="chartFornec"></canvas>
      </div>
    </div>
  </div>

  <!-- ── TABELAS ── -->
  <div class="section-header">
    <h2>📑 Transações Detalhadas</h2>
    <div class="section-divider"></div>
    <span class="section-badge">Dados brutos</span>
  </div>
  <div class="table-card">
    <div class="tab-btns">
      <button class="tab-btn active" onclick="switchTab('receitas')">💰 Receitas ({{ data.qtd_receitas }})</button>
      <button class="tab-btn" onclick="switchTab('despesas')">📋 Despesas ({{ data.qtd_despesas }})</button>
    </div>

    <!-- Receitas -->
    <div class="tab-panel active" id="tab-receitas">
      <div class="table-filters">
        <div>
          <div class="filter-label">Buscar</div>
          <input class="filter-input" id="filter-rec" type="text" placeholder="Nome, descrição, categoria..." oninput="filterTable('rec')">
        </div>
        <div>
          <div class="filter-label">Categoria</div>
          <select class="filter-input" id="filter-rec-cat" onchange="filterTable('rec')">
            <option value="">Todas</option>
            {% for cat in rec_cats %}
            <option value="{{ cat }}">{{ cat }}</option>
            {% endfor %}
          </select>
        </div>
        <div>
          <div class="filter-label">Banco</div>
          <select class="filter-input" id="filter-rec-banco" onchange="filterTable('rec')">
            <option value="">Todos</option>
            {% for b in rec_bancos %}
            <option value="{{ b }}">{{ b }}</option>
            {% endfor %}
          </select>
        </div>
      </div>
      <div class="table-wrapper">
        <table id="tbl-rec">
          <thead>
            <tr>
              <th>Data Pgto</th>
              <th>Nome / Descrição</th>
              <th>Categoria</th>
              <th>Banco</th>
              <th style="text-align:right">Valor</th>
            </tr>
          </thead>
          <tbody>
            {% for r in data.receitas_rows %}
            <tr>
              <td>{{ r.get('Data de pagamento') or r.get('Vencimento') or '—' }}</td>
              <td style="max-width:260px;overflow:hidden;text-overflow:ellipsis;">{{ r.get('Nome') or r.get('Descrição') or '—' }}</td>
              <td><span class="badge badge-blue">{{ r.get('Categoria') or '—' }}</span></td>
              <td>{{ (r.get('Banco') or '—')|trim }}</td>
              <td class="valor-pos" style="text-align:right">{{ fmt(r.get('Valor categoria/centro de custo') or 0) }}</td>
            </tr>
            {% endfor %}
          </tbody>
        </table>
      </div>
      <div class="pagination" id="pag-rec"></div>
    </div>

    <!-- Despesas -->
    <div class="tab-panel" id="tab-despesas">
      <div class="table-filters">
        <div>
          <div class="filter-label">Buscar</div>
          <input class="filter-input" id="filter-desp" type="text" placeholder="Nome, descrição, categoria..." oninput="filterTable('desp')">
        </div>
        <div>
          <div class="filter-label">Categoria</div>
          <select class="filter-input" id="filter-desp-cat" onchange="filterTable('desp')">
            <option value="">Todas</option>
            {% for cat in desp_cats %}
            <option value="{{ cat }}">{{ cat }}</option>
            {% endfor %}
          </select>
        </div>
        <div>
          <div class="filter-label">Banco</div>
          <select class="filter-input" id="filter-desp-banco" onchange="filterTable('desp')">
            <option value="">Todos</option>
            {% for b in desp_bancos %}
            <option value="{{ b }}">{{ b }}</option>
            {% endfor %}
          </select>
        </div>
      </div>
      <div class="table-wrapper">
        <table id="tbl-desp">
          <thead>
            <tr>
              <th>Data Pgto</th>
              <th>Nome / Beneficiário</th>
              <th>Categoria</th>
              <th>Banco</th>
              <th style="text-align:right">Valor</th>
            </tr>
          </thead>
          <tbody>
            {% for r in data.despesas_rows %}
            <tr>
              <td>{{ r.get('Data de pagamento') or r.get('Vencimento') or '—' }}</td>
              <td style="max-width:260px;overflow:hidden;text-overflow:ellipsis;">{{ r.get('Nome') or r.get('Descrição') or '—' }}</td>
              <td><span class="badge badge-red">{{ r.get('Categoria') or '—' }}</span></td>
              <td>{{ (r.get('Banco') or '—')|trim }}</td>
              <td class="valor-neg" style="text-align:right">{{ fmt(abs(r.get('Valor categoria/centro de custo') or 0)) }}</td>
            </tr>
            {% endfor %}
          </tbody>
        </table>
      </div>
      <div class="pagination" id="pag-desp"></div>
    </div>
  </div>

</div><!-- /container -->

<footer class="footer">
  <strong style="color:var(--blue)">IFP</strong> — Instituto de Formação Profissional · Dashboard Financeiro · Gerado em <span id="footer-date"></span>
</footer>

<script>
// ── Data atual ──
const now = new Date();
const dtStr = now.toLocaleDateString('pt-BR', { weekday:'long', year:'numeric', month:'long', day:'numeric' });
document.getElementById('header-date').textContent = dtStr;
document.getElementById('footer-date').textContent = now.toLocaleDateString('pt-BR');

// ── Dados do servidor ──
const D = {{ chart_data|safe }};

// ── Paleta ──
const BLUE_SHADES = [
  '#003087','#0047AB','#1565C0','#1976D2','#1E88E5',
  '#2196F3','#42A5F5','#64B5F6','#90CAF9','#BBDEFB',
];
const RED_SHADES = [
  '#CC0000','#E53935','#EF5350','#F44336','#EF9A9A',
  '#FFCDD2','#C62828','#B71C1C','#FF5252','#FF8A80',
];
const MULTI_COLORS = ['#003087','#CC0000','#0047AB','#E53935','#1565C0','#C62828','#1E88E5','#EF5350','#42A5F5','#FF5252','#64B5F6','#FFCDD2','#90CAF9','#1976D2','#2196F3'];

const baseChartOpts = {
  responsive: true, maintainAspectRatio: false,
  plugins: { legend: { labels: { font: { family: 'Segoe UI', size: 12 }, color: '#4A5568' } } }
};

// ── 1. Fluxo Diário ──
new Chart(document.getElementById('chartFluxo'), {
  type: 'bar',
  data: {
    labels: D.daily_labels,
    datasets: [
      {
        label: 'Receitas',
        data: D.daily_rec_vals,
        backgroundColor: 'rgba(0,48,135,0.75)',
        borderColor: '#003087',
        borderWidth: 1, borderRadius: 4,
      },
      {
        label: 'Despesas',
        data: D.daily_desp_vals,
        backgroundColor: 'rgba(204,0,0,0.7)',
        borderColor: '#CC0000',
        borderWidth: 1, borderRadius: 4,
      }
    ]
  },
  options: {
    ...baseChartOpts,
    interaction: { mode: 'index', intersect: false },
    scales: {
      x: { grid: { display: false }, ticks: { maxRotation: 45, font: { size: 11 } } },
      y: {
        grid: { color: 'rgba(0,0,0,0.05)' },
        ticks: { callback: v => 'R$ ' + v.toLocaleString('pt-BR', { minimumFractionDigits: 0 }) }
      }
    },
    plugins: { ...baseChartOpts.plugins, legend: { ...baseChartOpts.plugins.legend, position: 'top' } }
  }
});

// ── 2. Saldo Acumulado ──
new Chart(document.getElementById('chartSaldo'), {
  type: 'line',
  data: {
    labels: D.daily_labels,
    datasets: [{
      label: 'Saldo Acumulado',
      data: D.saldo_acum,
      borderColor: D.saldo_acum[D.saldo_acum.length-1] >= 0 ? '#0A7C3E' : '#CC0000',
      backgroundColor: D.saldo_acum[D.saldo_acum.length-1] >= 0
        ? 'rgba(10,124,62,0.08)' : 'rgba(204,0,0,0.08)',
      fill: true, tension: 0.35, pointRadius: 3,
      borderWidth: 2.5,
    }]
  },
  options: {
    ...baseChartOpts,
    scales: {
      x: { grid: { display: false }, ticks: { maxRotation: 45, font: { size: 11 } } },
      y: {
        grid: { color: 'rgba(0,0,0,0.05)' },
        ticks: { callback: v => 'R$ ' + v.toLocaleString('pt-BR', { minimumFractionDigits: 0 }) }
      }
    }
  }
});

// ── 3. Receitas por Categoria (Doughnut) ──
const rcCats = Object.keys(D.rec_por_cat);
const rcVals = Object.values(D.rec_por_cat);
new Chart(document.getElementById('chartRecCat'), {
  type: 'doughnut',
  data: {
    labels: rcCats,
    datasets: [{ data: rcVals, backgroundColor: BLUE_SHADES.slice(0, rcCats.length), borderWidth: 2, borderColor: '#fff' }]
  },
  options: {
    ...baseChartOpts,
    plugins: {
      legend: { position: 'right', labels: { font: { size: 11 }, padding: 12 } },
      tooltip: { callbacks: { label: ctx => ` R$ ${ctx.parsed.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}` } }
    }
  }
});

// ── 4. Despesas por Categoria (Horizontal Bar) ──
const dcCats = Object.keys(D.desp_por_cat);
const dcVals = Object.values(D.desp_por_cat);
new Chart(document.getElementById('chartDespCat'), {
  type: 'bar',
  data: {
    labels: dcCats,
    datasets: [{
      label: 'Valor (R$)',
      data: dcVals,
      backgroundColor: dcVals.map((_, i) => RED_SHADES[i % RED_SHADES.length]),
      borderRadius: 5, borderSkipped: false,
    }]
  },
  options: {
    ...baseChartOpts,
    indexAxis: 'y',
    scales: {
      x: { ticks: { callback: v => 'R$ ' + v.toLocaleString('pt-BR', { minimumFractionDigits: 0 }) }, grid: { color: 'rgba(0,0,0,0.04)' } },
      y: { ticks: { font: { size: 11 } }, grid: { display: false } }
    },
    plugins: { legend: { display: false } }
  }
});

// ── 5. Rec por Banco ──
const rbK = Object.keys(D.rec_por_banco);
const rbV = Object.values(D.rec_por_banco);
new Chart(document.getElementById('chartRecBanco'), {
  type: 'pie',
  data: {
    labels: rbK,
    datasets: [{ data: rbV, backgroundColor: BLUE_SHADES.slice(0, rbK.length), borderWidth: 2, borderColor: '#fff' }]
  },
  options: {
    ...baseChartOpts,
    plugins: {
      legend: { position: 'bottom', labels: { font: { size: 10 }, padding: 8, boxWidth: 12 } },
      tooltip: { callbacks: { label: ctx => ` R$ ${ctx.parsed.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}` } }
    }
  }
});

// ── 6. Desp por Banco ──
const dbK = Object.keys(D.desp_por_banco);
const dbV = Object.values(D.desp_por_banco);
new Chart(document.getElementById('chartDespBanco'), {
  type: 'bar',
  data: {
    labels: dbK,
    datasets: [{ label: 'Despesas', data: dbV, backgroundColor: RED_SHADES.slice(0, dbK.length), borderRadius: 6 }]
  },
  options: {
    ...baseChartOpts,
    plugins: { legend: { display: false } },
    scales: {
      x: { grid: { display: false }, ticks: { font: { size: 10 }, maxRotation: 30 } },
      y: { ticks: { callback: v => 'R$' + (v/1000).toFixed(0) + 'k' }, grid: { color: 'rgba(0,0,0,0.04)' } }
    }
  }
});

// ── 7. Meios de Pagamento ──
const mpK = Object.keys(D.metodos_pagamento);
const mpV = Object.values(D.metodos_pagamento);
new Chart(document.getElementById('chartMetodos'), {
  type: 'doughnut',
  data: {
    labels: mpK,
    datasets: [{ data: mpV, backgroundColor: MULTI_COLORS.slice(0, mpK.length), borderWidth: 2, borderColor: '#fff' }]
  },
  options: {
    ...baseChartOpts,
    plugins: {
      legend: { position: 'bottom', labels: { font: { size: 10 }, padding: 8, boxWidth: 12 } },
      tooltip: { callbacks: { label: ctx => ` R$ ${ctx.parsed.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}` } }
    }
  }
});

// ── 8. Top Fornecedores (Bar + Lista) ──
const fK = D.top_fornecedores.map(f => f[0]);
const fV = D.top_fornecedores.map(f => f[1]);
const maxF = Math.max(...fV);

new Chart(document.getElementById('chartFornec'), {
  type: 'doughnut',
  data: {
    labels: fK,
    datasets: [{ data: fV, backgroundColor: RED_SHADES.slice(0, fK.length), borderWidth: 2, borderColor: '#fff' }]
  },
  options: {
    ...baseChartOpts,
    plugins: {
      legend: { position: 'bottom', labels: { font: { size: 10 }, padding: 6, boxWidth: 10 } },
      tooltip: { callbacks: { label: ctx => ` R$ ${ctx.parsed.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}` } }
    }
  }
});

const fornecList = document.getElementById('fornec-list');
D.top_fornecedores.forEach(([nome, val], i) => {
  const pct = Math.round(val / maxF * 100);
  const div = document.createElement('div');
  div.className = 'fornec-item';
  const shortNome = nome.length > 34 ? nome.slice(0, 32) + '…' : nome;
  div.innerHTML = `
    <div class="fornec-rank ${i<3?'top3':''}">${i+1}</div>
    <div style="flex:1;min-width:0">
      <div class="fornec-name" title="${nome}">${shortNome}</div>
      <div class="fornec-bar"><div class="fornec-bar-fill" style="width:${pct}%"></div></div>
    </div>
    <div class="fornec-val">R$ ${val.toLocaleString('pt-BR',{minimumFractionDigits:2})}</div>
  `;
  fornecList.appendChild(div);
});

// ── TABS ──
function switchTab(tab) {
  document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
  document.getElementById('tab-' + tab).classList.add('active');
  event.target.classList.add('active');
}

// ── FILTRO DE TABELA ──
function filterTable(which) {
  const q     = (document.getElementById('filter-' + which).value || '').toLowerCase();
  const cat   = (document.getElementById('filter-' + which + '-cat').value || '').toLowerCase();
  const banco = (document.getElementById('filter-' + which + '-banco').value || '').toLowerCase();
  const tbl   = document.getElementById('tbl-' + which);
  const rows  = tbl.querySelectorAll('tbody tr');
  rows.forEach(row => {
    const text  = row.textContent.toLowerCase();
    const cells = row.querySelectorAll('td');
    const rowCat   = (cells[2]?.textContent || '').toLowerCase();
    const rowBanco = (cells[3]?.textContent || '').toLowerCase();
    const show = (!q || text.includes(q))
              && (!cat   || rowCat.includes(cat))
              && (!banco || rowBanco.includes(banco));
    row.style.display = show ? '' : 'none';
  });
}
</script>
</body>
</html>
"""

# ─────────────────────────────────────────────────────────────
# ROTAS
# ─────────────────────────────────────────────────────────────
@app.route("/")
def index():
    messages = session.pop("flash_messages", [])
    return render_template_string(UPLOAD_TEMPLATE, messages=messages)


@app.route("/upload", methods=["POST"])
def upload():
    rec_file  = request.files.get("receitas")
    desp_file = request.files.get("despesas")

    if not rec_file or not desp_file:
        session["flash_messages"] = [("Por favor, envie os dois arquivos.", "error")]
        return redirect(url_for("index"))

    if not rec_file.filename.endswith(".xlsx") or not desp_file.filename.endswith(".xlsx"):
        session["flash_messages"] = [("Os arquivos devem estar no formato .xlsx", "error")]
        return redirect(url_for("index"))

    try:
        uid = str(uuid.uuid4())[:8]
        rec_path  = os.path.join(UPLOAD_FOLDER, f"{uid}_receitas.xlsx")
        desp_path = os.path.join(UPLOAD_FOLDER, f"{uid}_despesas.xlsx")
        rec_file.save(rec_path)
        desp_file.save(desp_path)

        _, rec_rows  = read_excel(rec_path)
        _, desp_rows = read_excel(desp_path)

        session["rec_path"]  = rec_path
        session["desp_path"] = desp_path
        session["processed"] = True

        return redirect(url_for("dashboard"))
    except Exception as e:
        session["flash_messages"] = [(f"Erro ao processar arquivos: {str(e)}", "error")]
        return redirect(url_for("index"))


@app.route("/dashboard")
def dashboard():
    rec_path  = session.get("rec_path")
    desp_path = session.get("desp_path")

    if not rec_path or not desp_path:
        return redirect(url_for("index"))

    try:
        _, rec_rows  = read_excel(rec_path)
        _, desp_rows = read_excel(desp_path)
    except Exception:
        return redirect(url_for("index"))

    data = parse_data(desp_rows, rec_rows)

    # Serializa dados para Chart.js
    chart_data = {
        "daily_labels":    data["daily_labels"],
        "daily_rec_vals":  data["daily_rec_vals"],
        "daily_desp_vals": data["daily_desp_vals"],
        "saldo_acum":      data["saldo_acum"],
        "rec_por_cat":     data["rec_por_cat"],
        "desp_por_cat":    data["desp_por_cat"],
        "rec_por_banco":   data["rec_por_banco"],
        "desp_por_banco":  data["desp_por_banco"],
        "metodos_pagamento": data["metodos_pagamento"],
        "top_fornecedores": data["top_fornecedores"],
    }

    # Listas de filtros únicos
    rec_cats   = sorted(set(r.get("Categoria") or "" for r in rec_rows  if r.get("Categoria")))
    desp_cats  = sorted(set(r.get("Categoria") or "" for r in desp_rows if r.get("Categoria")))
    rec_bancos = sorted(set((r.get("Banco") or "").strip() for r in rec_rows  if r.get("Banco")))
    desp_bancos= sorted(set((r.get("Banco") or "").strip() for r in desp_rows if r.get("Banco")))

    def fmt(val):
        try:
            v = float(val or 0)
            formatted = f"{abs(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            return f"R$ {formatted}"
        except Exception:
            return "R$ 0,00"

    return render_template_string(
        DASHBOARD_TEMPLATE,
        data=data,
        chart_data=json.dumps(chart_data),
        rec_cats=rec_cats,
        desp_cats=desp_cats,
        rec_bancos=rec_bancos,
        desp_bancos=desp_bancos,
        fmt=fmt,
    )


@app.route("/api/data")
def api_data():
    """Endpoint JSON para integrações externas."""
    rec_path  = session.get("rec_path")
    desp_path = session.get("desp_path")
    if not rec_path or not desp_path:
        return jsonify({"error": "No data loaded"}), 404
    _, rec_rows  = read_excel(rec_path)
    _, desp_rows = read_excel(desp_path)
    data = parse_data(desp_rows, rec_rows)
    return jsonify({
        "total_receita": data["total_receita"],
        "total_despesa": data["total_despesa"],
        "resultado":     data["resultado"],
        "margem":        data["margem"],
    })


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
