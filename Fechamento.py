import os
import re
import pandas as pd
import pdfplumber
from collections import Counter


def formatar_meio_de_pagamento(observacoes):
    """
    Analisa as observações de uma venda para identificar e formatar os meios de pagamento.
    """
    pagamentos_encontrados = []

    # Padrão para pagamentos múltiplos (ex: R$ 125,00 no dinheiro)
    padrao_multiplo = re.findall(r'R\$\s*([\d.,]+)\s*.*?(dinheiro|master|elo|pix)', observacoes, re.IGNORECASE)

    if padrao_multiplo:
        for valor, tipo in padrao_multiplo:
            valor_formatado = valor.strip().replace('.', '').replace(',', '.')
            if 'dinheiro' in tipo.lower():
                pagamentos_encontrados.append(f"Dinheiro (R$ {float(valor_formatado):.2f})")
            elif 'master' in tipo.lower():
                pagamentos_encontrados.append(f"Cartão de Crédito (R$ {float(valor_formatado):.2f})")
            elif 'elo' in tipo.lower():
                pagamentos_encontrados.append(f"Cartão de Débito (R$ {float(valor_formatado):.2f})")
            elif 'pix' in tipo.lower():
                pagamentos_encontrados.append(f"PIX (R$ {float(valor_formatado):.2f})")

    if pagamentos_encontrados:
        return ", ".join(pagamentos_encontrados)

    # Lógica para pagamentos únicos baseada em palavras-chave
    if 'link de pagamento' in observacoes.lower(): return 'Cartão de Crédito'
    if 'elo' in observacoes.lower(): return 'Cartão de Débito'
    if 'a receber' in observacoes.lower() or 'pix' in observacoes.lower(): return 'Transferência/PIX'
    if 'dinheiro' in observacoes.lower(): return 'Dinheiro'

    return 'Não Especificado'


def encontrar_pagamento_mais_frequente(df_vendas):
    """
    Calcula o meio de pagamento mais comum a partir da coluna de pagamentos.
    """
    if df_vendas.empty: return "Nenhum"

    lista_pagamentos = []
    for item in df_vendas['Meio de Pagamento']:
        tipos = re.findall(r'([a-zA-Zãáçéíõú\s/]+)', item)
        for tipo in tipos:
            tipo_limpo = tipo.strip().replace('(', '').replace(')', '').strip()
            if tipo_limpo and 'R$' not in tipo_limpo:
                lista_pagamentos.append(tipo_limpo)

    if not lista_pagamentos: return "Nenhum"
    return Counter(lista_pagamentos).most_common(1)[0][0]


def extrair_dados_dos_pdfs(pasta_dos_pdfs):
    """
    Lê todos os PDFs em uma pasta e extrai as informações das transações.
    """
    lista_de_transacoes = []

    if not os.path.isdir(pasta_dos_pdfs):
        print(
            f"ERRO: A pasta '{pasta_dos_pdfs}' não foi encontrada. Por favor, crie a pasta e coloque seus relatórios dentro dela.")
        return None

    for nome_arquivo in os.listdir(pasta_dos_pdfs):
        if nome_arquivo.lower().endswith('.pdf'):
            caminho_completo = os.path.join(pasta_dos_pdfs, nome_arquivo)
            print(f"Processando o arquivo: {nome_arquivo}...")

            try:
                with pdfplumber.open(caminho_completo) as pdf:
                    texto_completo = "".join([page.extract_text() + "\n" for page in pdf.pages if page.extract_text()])
                    blocos_transacao = re.split(r'\n(?=\d{4}\s)', texto_completo)

                    for bloco in blocos_transacao:
                        padrao_transacao = re.search(
                            r'(\d{4})\s+(\d{2}/\d{2}/\d{4})\s+[\d:]+\s+(.*?)\s+R\$\s[\d.,]+\s+R\$\s([\d.,]+)([\s\S]*)',
                            bloco, re.DOTALL
                        )
                        if padrao_transacao:
                            cliente_bruto = padrao_transacao.group(3).strip()
                            if 'abertura de caixa' in cliente_bruto.lower(): continue

                            num_pedido = padrao_transacao.group(1).strip()
                            data_venda = padrao_transacao.group(2).strip()
                            valor_venda_str = padrao_transacao.group(4).strip().replace('.', '').replace(',', '.')
                            observacoes = padrao_transacao.group(5).strip()

                            lista_de_transacoes.append({
                                'Numero do Pedido': num_pedido,
                                'Data': data_venda,
                                'Nome do Cliente': cliente_bruto,
                                'Valor Total': float(valor_venda_str),
                                'Meio de Pagamento': formatar_meio_de_pagamento(observacoes)
                            })
            except Exception as e:
                print(f"Erro ao processar o arquivo {nome_arquivo}: {e}")
    return lista_de_transacoes


def criar_planilha_excel(dados_vendas, caminho_saida):
    """
    Cria a planilha Excel com os dados das vendas, ordenados e com resumo final.
    """
    ordem_colunas = ['Numero do Pedido', 'Data', 'Nome do Cliente', 'Valor Total', 'Meio de Pagamento']
    df = pd.DataFrame(dados_vendas)

    # Converte as colunas para o tipo correto antes de ordenar
    df['Numero do Pedido'] = pd.to_numeric(df['Numero do Pedido'])
    df['Data'] = pd.to_datetime(df['Data'], format='%d/%m/%Y')

    # **NOVO: Ordena os dados pelo número do pedido e depois pela data**
    df = df.sort_values(by=['Numero do Pedido', 'Data'])

    # Formata a data de volta para o formato dia/mês/ano para exibição no Excel
    df['Data'] = df['Data'].dt.strftime('%d/%m/%Y')

    # Garante a ordem final das colunas
    df = df[ordem_colunas]

    total_mes = df['Valor Total'].sum()
    pagamento_frequente = encontrar_pagamento_mais_frequente(df)

    with pd.ExcelWriter(caminho_saida, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Relatorio_Mensal')

        workbook = writer.book
        worksheet = writer.sheets['Relatorio_Mensal']

        linha_resumo = len(df) + 3
        worksheet.cell(row=linha_resumo, column=1, value='Total de Vendas no Mês:')
        worksheet.cell(row=linha_resumo, column=2, value=total_mes).number_format = '"R$" #,##0.00'

        worksheet.cell(row=linha_resumo + 1, column=1, value='Meio de Pagamento Mais Frequente:')
        worksheet.cell(row=linha_resumo + 1, column=2, value=pagamento_frequente)

        for column_cells in worksheet.columns:
            max_length = max(len(str(cell.value)) for cell in column_cells)
            worksheet.column_dimensions[column_cells[0].column_letter].width = max_length + 2

    print(f"\nPlanilha '{caminho_saida}' gerada com sucesso e em ordem!")


# --- INÍCIO DA EXECUÇÃO ---
if __name__ == "__main__":
    pasta_dos_pdfs = "PDF"

    dados_extraidos = extrair_dados_dos_pdfs(pasta_dos_pdfs)

    if dados_extraidos:
        nome_arquivo_excel = "Relatorio_Mensal_Completo.xlsx"
        criar_planilha_excel(dados_extraidos, nome_arquivo_excel)
    elif dados_extraidos is None:
        pass
    else:
        print("Nenhuma transação de venda foi encontrada nos arquivos PDF.")