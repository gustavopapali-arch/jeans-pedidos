from flask import Flask, render_template, request, send_file, redirect, session
import os
from openpyxl import Workbook

app = Flask(__name__)
app.secret_key = "SENHA_SEGURA"

TAMANHOS = ["34","36","38","40","42","44","46"]

@app.route("/")
def index():
    return render_template("index.html", tamanhos=TAMANHOS)

@app.route("/enviar", methods=["POST"])
def enviar():
    nome = request.form.get("nome")
    cor = request.form.get("cor")
    quantidade = request.form.get("quantidade")
    tamanho = request.form.get("tamanho")

    wb = Workbook()
    ws = wb.active
    ws.append(["Nome","Cor","Quantidade","Tamanho"])
    ws.append([nome, cor, quantidade, tamanho])

    if not os.path.exists("pedidos"):
        os.makedirs("pedidos")

    filename = f"pedido_{nome}.xlsx"
    wb.save(os.path.join("pedidos", filename))
    return "Pedido enviado com sucesso!"

@app.route("/login", methods=["GET","POST"])
def login():
    if request.method == "POST":
        if request.form.get("senha") == "1234":
            session["admin"] = True
            return redirect("/admin")
        return "Senha incorreta"
    return '<form method="POST"><input type="password" name="senha" placeholder="Senha"/><button type="submit">Entrar</button></form>'

@app.route("/admin")
def admin():
    if not session.get("admin"):
        return redirect("/login")

    arquivos = sorted(os.listdir("pedidos")) if os.path.exists("pedidos") else []

    return render_template("admin.html", pedidos=arquivos)

@app.route("/download/<arquivo>")
def download(arquivo):
    caminho = os.path.join("pedidos", arquivo)
    return send_file(caminho, as_attachment=True)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
