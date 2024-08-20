from flask import Flask, request, jsonify
import openpyxl

app = Flask(__name__)

excel_path = "vendas.xlsx"  # Caminho da planilha

@app.route('/save_product', methods=['POST'])
def save_product():
    data = request.get_json()
    product = data['product']
    quantity = data['quantity']
    stock = data['stock']

    # Carregar a planilha do Excel
    try:
        wb = openpyxl.load_workbook(excel_path)
    except FileNotFoundError:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Vendas"
        ws.append(["Produto", "Quantidade", "Estoque Disponível"])

    ws = wb.active

    # Adicionar os dados na planilha
    ws.append([product, quantity, stock])

    wb.save(excel_path)
    return jsonify({"status": "success"})

if __name__ == '__main__':
    app.run(debug=True)
