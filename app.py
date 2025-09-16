from flask import Flask, render_template, request, send_file, make_response
import io
import openpyxl

from modelo1.parte1_processador import executar_processo_parte1
from modelo1.parte2_processador import executar_comparacao_lado_a_lado, executar_comparacao_com_exclusao_parcial
from modelo1.resumo_processador import criar_aba_resumo

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process():
    if 'excel_file' not in request.files:
        return "Erro: Nenhum ficheiro foi enviado.", 400
    
    uploaded_file = request.files['excel_file']
    modelo_selecionado = request.form.get('modelo_selecionado')
    
    data_corte_compras = request.form.get('data_corte_compras')
    data_corte_vendas = request.form.get('data_corte_vendas')

    if uploaded_file.filename == '':
        return "Erro: Nenhum ficheiro foi selecionado.", 400
        
    if not modelo_selecionado:
        return "Erro: Por favor, selecione um modelo de folha de cálculo.", 400

    if uploaded_file:
        try:
            print(f"Ficheiro recebido. A processar com o '{modelo_selecionado}'...")
            
            workbook_resultado = None

            if modelo_selecionado == 'modelo1':
                p1_pular_linhas_antes_cabecalho = 0
                p1_nomes_colunas = {
                    "nome": "Parte - Contra Banco", "cnpj": "CNPJ", "valor": "Valor Ajustado",
                    "deal": "Deal", "num_contrato": "Nº Contrato", "tipo_operacao": "Tipo de Operação",
                    "liquidacao": "Liquidação "
                }
                
                nomes_colunas_saida_unico = ['CNPJ Esquerdo', 'Nota', 'Data Emissão', 'Fornecedor', 'Valor Contábil', ' ', 'CNPJ Direito', 'Nº Contrato', 'Liquidação', 'Parte - Contra Banco', 'Valor Ajustado', 'Diferença do Bloco']
                nomes_colunas_saida_cliente = ['CNPJ Esquerdo', 'Nota', 'Data Emissão', 'Cliente', 'Valor Contábil', ' ', 'CNPJ Direito', 'Nº Contrato', 'Liquidação', 'Parte - Contra Banco', 'Valor Ajustado', 'Diferença do Bloco']
                
                config_2_vs_5 = {
                    "nome_processo": "Parte 2.1 - Confronto Padrão", "arquivo_excel": uploaded_file, "indice_aba_1": 1, "indice_aba_2": 4,
                    "pular_linhas_1": 5, "pular_linhas_2": 0, 
                    "nome_aba_saida": "Livro x Book Entrada - CNPJ",
                    "colunas_aba_1": {"nome": "Fornecedor", "cnpj": "CNPJ/CPF/CEI/CAEPF", "valor": "Valor Contábil", "nota": "Nota", "data_emissao": "Data Emissão"},
                    "colunas_aba_2": {"nome": "Parte - Contra Banco", "cnpj": "CNPJ", "valor": "Valor Ajustado", "num_contrato": "Nº Contrato", "liquidacao": "Liquidação "},
                    "nomes_colunas_saida": nomes_colunas_saida_unico,
                    "chave_agrupamento_final": 12,
                    "data_corte": data_corte_compras
                }
                config_1_vs_6 = {
                    "nome_processo": "Parte 2.2 - Confronto Padrão", "arquivo_excel": uploaded_file, "indice_aba_1": 0, "indice_aba_2": 5,
                    "pular_linhas_1": 5, "pular_linhas_2": 0, 
                    "nome_aba_saida": "Livro x Book Saída - CNPJ",
                    "colunas_aba_1": {"nome": "Cliente", "cnpj": "CNPJ/CPF/CEI/CAEPF", "valor": "Valor Contábil", "nota": "Nota", "data_emissao": "Data Emissão"},
                    "colunas_aba_2": {"nome": "Parte - Contra Banco", "cnpj": "CNPJ", "valor": "Valor Ajustado", "num_contrato": "Nº Contrato", "liquidacao": "Liquidação "},
                    "nomes_colunas_saida": nomes_colunas_saida_cliente,
                    "chave_agrupamento_final": 12,
                    "data_corte": data_corte_vendas
                }
                
                config_parcial_2_vs_5 = {**config_2_vs_5, 
                    "nome_processo": "Parte 3.1 - Confronto com Exclusão", 
                    "nome_aba_saida": "Livro x Book Entrada",
                    "chave_agrupamento_final": 8
                }
                config_parcial_1_vs_6 = {**config_1_vs_6, 
                    "nome_processo": "Parte 3.2 - Confronto com Exclusão", 
                    "nome_aba_saida": "Livro x Book Saída",
                    "chave_agrupamento_final": 8
                }

                workbook_resultado = executar_processo_parte1(uploaded_file, p1_pular_linhas_antes_cabecalho, p1_nomes_colunas)
                workbook_resultado = executar_comparacao_lado_a_lado(workbook_resultado, config_2_vs_5)
                workbook_resultado = executar_comparacao_lado_a_lado(workbook_resultado, config_1_vs_6)
                workbook_resultado = executar_comparacao_com_exclusao_parcial(workbook_resultado, config_parcial_2_vs_5)
                workbook_resultado = executar_comparacao_com_exclusao_parcial(workbook_resultado, config_parcial_1_vs_6)
                
                # --- ETAPA FINAL: CRIAR ABA DE RESUMO ---
                if workbook_resultado:
                    workbook_resultado = criar_aba_resumo(
                        workbook_resultado,
                        p1_nomes_colunas,
                        data_corte_compras,
                        data_corte_vendas
                    )
            
            else:
                return f"Erro: Modelo '{modelo_selecionado}' não reconhecido.", 400

            if workbook_resultado:
                print("Processamento concluído. A preparar ficheiro para download...")
                output = io.BytesIO()
                workbook_resultado.save(output)
                output.seek(0)
                
                nome_arquivo_saida = uploaded_file.filename

                response = make_response(send_file(
                    output, as_attachment=True, download_name=nome_arquivo_saida,
                    mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                ))
                response.set_cookie('fileDownload', 'true', max_age=60)
                return response
            else:
                return "Ocorreu um erro durante o processamento.", 500
        except Exception as e:
            print(f"Erro detalhado: {e}")
            return f"Ocorreu um erro inesperado durante o processamento: {e}", 500
    return "Erro desconhecido.", 500

if __name__ == '__main__':
    app.run(debug=True, port=8080)

