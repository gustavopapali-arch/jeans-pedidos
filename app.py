from flask import Flask, render_template, request, redirect, url_for, send_file
import pandas as pd
from datetime import datetime
import os

app = Flask(__name__)
app.secret_key = os.environ.get("SESSION_SECRET", "dev-secret-key")

EXCEL_FILE = 'pedidos.xlsx'
TAMANHOS = [36, 38, 40, 42, 44, 46, 48, 50, 52, 54, 56]
SECRET_URL = 'secret123'


def carregar_pedidos():
    """Carrega pedidos do arquivo Excel usando pandas"""
    if not os.path.exists(EXCEL_FILE):
        return pd.DataFrame()
    
    try:
        df = pd.read_excel(EXCEL_FILE, engine='openpyxl')
        return df
    except Exception:
        return pd.DataFrame()


def adicionar_pedido(nome_cliente, quantidades_azul, quantidades_preta):
    """Adiciona um novo pedido ao arquivo Excel"""
    timestamp = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
    
    novo_pedido = {
        'Timestamp': timestamp,
        'Nome do Cliente': nome_cliente
    }
    
    for tamanho in TAMANHOS:
        novo_pedido[f'Azul_{tamanho}'] = quantidades_azul.get(str(tamanho), 0)
        novo_pedido[f'Preta_{tamanho}'] = quantidades_preta.get(str(tamanho), 0)
    
    if os.path.exists(EXCEL_FILE):
        df = carregar_pedidos()
        novo_df = pd.DataFrame([novo_pedido])
        df = pd.concat([df, novo_df], ignore_index=True)
    else:
        df = pd.DataFrame([novo_pedido])
    
    df.to_excel(EXCEL_FILE, index=False, engine='openpyxl')


def calcular_totais(df):
    """Calcula totais de peças azuis e pretas"""
    if df.empty:
        return 0, 0
    
    total_azul = 0
    total_preta = 0
    
    for tamanho in TAMANHOS:
        col_azul = f'Azul_{tamanho}'
        col_preta = f'Preta_{tamanho}'
        
        if col_azul in df.columns:
            total_azul += df[col_azul].sum()
        if col_preta in df.columns:
            total_preta += df[col_preta].sum()
    
    return int(total_azul), int(total_preta)


def preparar_lista_pedidos(df):
    """Prepara lista de pedidos para exibição"""
    if df.empty:
        return []
    
    pedidos = []
    for _, row in df.iterrows():
        total_azul = sum(row.get(f'Azul_{t}', 0) or 0 for t in TAMANHOS)
        total_preta = sum(row.get(f'Preta_{t}', 0) or 0 for t in TAMANHOS)
        
        pedidos.append({
            'timestamp': row.get('Timestamp', ''),
            'nome': row.get('Nome do Cliente', ''),
            'total_azul': int(total_azul),
            'total_preta': int(total_preta),
            'total_geral': int(total_azul + total_preta)
        })
    
    pedidos.sort(
        key=lambda x: datetime.strptime(
            x['timestamp'], '%d/%m/%Y %H:%M:%S'
        ) if x['timestamp'] else datetime.min,
        reverse=True
    )
    
    return pedidos


@app.route('/')
def index():
    """Página do formulário de pedidos"""
    return render_template('index.html', tamanhos=TAMANHOS)


@app.route('/enviar', methods=['POST'])
def enviar_pedido():
    """Processa o envio do pedido"""
    nome_cliente = request.form.get('nome', '').strip()
    
    if not nome_cliente:
        return redirect(url_for('index'))
    
    quantidades_azul = {}
    quantidades_preta = {}
    
    for tamanho in TAMANHOS:
        qtd_azul = request.form.get(f'azul_{tamanho}', '0')
        qtd_preta = request.form.get(f'preta_{tamanho}', '0')
        
        try:
            quantidades_azul[str(tamanho)] = int(qtd_azul) if qtd_azul else 0
            quantidades_preta[str(tamanho)] = int(qtd_preta) if qtd_preta else 0
        except ValueError:
            quantidades_azul[str(tamanho)] = 0
            quantidades_preta[str(tamanho)] = 0
    
    adicionar_pedido(nome_cliente, quantidades_azul, quantidades_preta)
    
    return redirect(url_for('index', sucesso='1'))


@app.route(f'/empresa/{SECRET_URL}')
def empresa():
    """Painel administrativo da empresa (restrito por URL)"""
    df = carregar_pedidos()
    total_azul, total_preta = calcular_totais(df)
    pedidos = preparar_lista_pedidos(df)
    total_pedidos = len(pedidos)
    
    return render_template('empresa.html',
                         pedidos=pedidos,
                         total_azul=total_azul,
                         total_preta=total_preta,
                         total_pedidos=total_pedidos,
                         secret_url=SECRET_URL)


@app.route(f'/empresa/{SECRET_URL}/download')
def download_excel():
    """Download do arquivo Excel de pedidos"""
    if os.path.exists(EXCEL_FILE):
        return send_file(EXCEL_FILE, as_attachment=True, download_name='pedidos.xlsx')
    return redirect(url_for('empresa'))


# ============================================================
# ➜ ADIÇÃO: ÁREA DA EMPRESA (NOVA ROTA)
# ============================================================

@app.route('/empresa-area')
def empresa_area():
    """
    Nova área da empresa
    Mantém toda a estrutura do sistema
    Não interfere com nada existente
    """
    df = carregar_pedidos()
    total_azul, total_preta = calcular_totais(df)
    pedidos = preparar_lista_pedidos(df)

    return render_template(
        'empresa_area.html',
        pedidos=pedidos,
        total_azul=total_azul,
        total_preta=total_preta,
        total_pedidos=len(pedidos)
    )


# ============================================================

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
