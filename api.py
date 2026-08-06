from flask import Flask, request, jsonify
from orcador_pecas import CatalogManager, TaxEngine, ItemOrcamento
import os

app = Flask(__name__)

# Carrega o catálogo e o motor fiscal usando a sua planilha
caminho_planilha = "LISTA DE PREÇO PEÇAS_050526.xlsx"
catalog = CatalogManager(caminho_planilha)
engine = TaxEngine(catalog.df_fiscal)

@app.route('/', methods=['GET'])
def home():
    return jsonify({"status": "API do Sistema Effa Rodando com Sucesso!"})

@app.route('/calcular', methods=['POST'])
def calcular():
    # Recebe o JSON que o AppSheet vai enviar
    dados = request.json
    uf_destino = dados.get('uf_destino', 'SP')
    canal = dados.get('canal', 'REVENDAS/OFICINAS')
    frete_total = float(dados.get('valor_frete_total', 0.0))
    itens_recebidos = dados.get('itens', [])
    
    total_itens_qtd = sum(i['quantidade'] for i in itens_recebidos)
    frete_unitario = frete_total
    
    resultados = []
    tot_geral = 0.0
    
    # Roda o motor fiscal para cada peça que o AppSheet mandou
    for item_dict in itens_recebidos:
        item_obj = ItemOrcamento(
            codigo=item_dict['codigo'],
            descricao=item_dict.get('nome', ''),
            ncm=item_dict.get('ncm', ''),
            quantidade=item_dict['quantidade'],
            preco_base=float(item_dict['preco_base']),
            ipi_percent=float(item_dict.get('ipi_percent', 0.0)),
            canal=canal
        )
        res = engine.calcular_item(item_obj, uf_destino=uf_destino, frete_unitario=frete_unitario)
        tot_geral += res.preco_final_total
        
        resultados.append({
            "codigo": res.item.codigo,
            "quantidade": res.item.quantidade,
            "valor_ipi": res.valor_ipi,
            "valor_icms_st": res.valor_icms_st,
            "preco_final_unitario": res.preco_final_unitario,
            "preco_final_total": res.preco_final_total
        })
        
    # Devolve a resposta pronta para o AppSheet
    return jsonify({
        "status": "sucesso",
        "total_geral": round(tot_geral, 2),
        "itens_calculados": resultados
    })

if __name__ == '__main__':
    # Roda o servidor na porta 5000
    app.run(host='0.0.0.0', port=5000)