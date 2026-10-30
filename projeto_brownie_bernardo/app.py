from flask import Flask, render_template

app = Flask(__name__)

@app.route('/')
def home():
    lista_produtos = [
       {"nome": "Brownie Tradicional", "descricao": "Massa molhadinha com casquinha crocante", "preco": 8.00},
        {"nome": "Brownie com Doce de Leite", "descricao": "Recheado com doce de leite artesanal", "preco": 10.00},
        {"nome": "Brownie com Ninho", "descricao": "Cobertura generosa de leite Ninho", "preco": 11.50}
    ]

    return render_template('cardapio.html', produtos = lista_produtos)

import os

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)