import fitz  # PyMuPDF
import re
import os
import pandas as pd
from typing import List, Dict

def extrair_dados_pdf(caminho_pdf: str) -> List[Dict[str, any]]:
    """
    Extrai dados de pedidos de um arquivo PDF com uma nova abordagem,
    processando o texto linha por linha para maior robustez.

    Args:
        caminho_pdf: O caminho completo para o arquivo PDF.

    Returns:
        Uma lista de dicionários, onde cada dicionário representa um item do pedido.
    """
    try:
        with fitz.open(caminho_pdf) as doc:
            # Extrai o texto página por página, mantendo uma certa ordem.
            texto_completo = "\n".join([page.get_text("text", sort=True) for page in doc])

        # --- INÍCIO DA DEPURAÇÃO ---
        print("\n" + "="*80)
        print(f"--- TEXTO EXTRAÍDO DE: {os.path.basename(caminho_pdf)} (visão linha a linha) ---")
        if not texto_completo.strip():
            print("AVISO: O PDF parece estar vazio ou contém apenas imagens. Nenhum texto foi extraído.")
        else:
            # Imprime cada linha para análise
            for i, linha in enumerate(texto_completo.split('\n')):
                print(f"Linha {i:03d}: {linha}")
        print("--- FIM DO TEXTO EXTRAÍDO ---")
        print("="*80 + "\n")
        # --- FIM DA DEPURAÇÃO ---

        # Extração de metadados (Cliente e Data de Entrega)
        match_cliente = re.search(r"CLIENTE:\s*\d*\s*-?\s*(.+)", texto_completo, re.IGNORECASE)
        cliente = match_cliente.group(1).strip() if match_cliente else "N/A"

        match_data = re.search(r"Previsão de entrega.*?(\d{2}/\d{2}/\d{4})", texto_completo, re.IGNORECASE) or \
                     re.search(r"Entrega:\s*(\d{2}/\d{2}/\d{4})", texto_completo, re.IGNORECASE)
        data_entrega = match_data.group(1) if match_data else "N/A"
        
        itens_encontrados = []
        # Processa o texto linha por linha
        for linha in texto_completo.split('\n'):
            # Verifica se a linha provavelmente contém um item do pedido
            # Condições: deve conter 'V-' ou 'V ', 'mm' e 'kg'
            if re.search(r'V[-\s]', linha, re.IGNORECASE) and 'mm' in linha.lower() and 'kg' in linha.lower():
                print(f"[DEBUG] Possível item encontrado na linha: '{linha}'")
                try:
                    # Tenta extrair cada parte da informação da linha
                    vergalhao = (re.search(r'(V[-\s]?[\d,]+)', linha, re.IGNORECASE) or ["N/A"])[0].replace(" ", "-")
                    material = (re.search(r'FIO DE\s+([A-ZÇÃÕÊÉÍ]+)', linha, re.IGNORECASE) or ["N/A", "N/A"])[1]
                    medidas = (re.search(r'([\d,]+\s*[xX]\s*[\d,]+)\s*mm', linha, re.IGNORECASE) or ["N/A"])[0]
                    peso_str = (re.search(r'([\d.,]+)\s*kg', linha, re.IGNORECASE) or ["0"])[0]
                    
                    # Limpa e formata os dados extraídos
                    diâmetro, comprimento = [m.strip() for m in medidas.replace(",", ".").split('x')] if 'x' in medidas.lower() else ["0", "0"]
                    peso = float(peso_str.replace(".", "").replace(",", "."))

                    item = {
                        "Cliente": cliente,
                        "Data_Entrega": data_entrega,
                        "Vergalhão": vergalhao,
                        "Material": f"FIO DE {material}",
                        "Diâmetro": diâmetro,
                        "Comprimento": comprimento,
                        "Peso": peso,
                        "Arquivo": os.path.basename(caminho_pdf)
                    }
                    itens_encontrados.append(item)
                except Exception as e:
                    print(f"⚠️  Aviso: Falha ao extrair dados da linha candidata. Linha: '{linha}'. Erro: {e}")
        
        return itens_encontrados

    except Exception as e:
        print(f"❌ Erro fatal ao abrir ou ler o arquivo PDF '{caminho_pdf}': {e}")
        return []


def gerar_excel(dados: List[Dict], nome_arquivo: str = "pedidos_extraidos.xlsx"):
    """Gera um arquivo Excel a partir de uma lista de dados extraídos."""
    if not dados:
        print("⚠️  Nenhum dado foi extraído para exportar para o Excel.")
        return

    df = pd.DataFrame(dados)
    
    colunas_finais = [
        "Cliente", "Data_Entrega", "Vergalhão", "Material", 
        "Diâmetro", "Comprimento", "Peso", "Arquivo"
    ]
    colunas_presentes = [col for col in colunas_finais if col in df.columns]
    df = df[colunas_presentes]
    
    df.rename(columns={
        'Data_Entrega': 'Data Entrega',
        'Diâmetro': 'Diâmetro (mm)',
        'Comprimento': 'Comprimento (mm)',
        'Peso': 'Peso (kg)'
    }, inplace=True)

    df.to_excel(nome_arquivo, index=False, engine='openpyxl')
    print(f"✅ Arquivo Excel '{nome_arquivo}' gerado com sucesso!")


def processar_pasta(pasta_pdfs: str = "pedidos"):
    """Processa todos os arquivos PDF em uma pasta especificada."""
    if not os.path.exists(pasta_pdfs):
        print(f"❌ Erro: A pasta '{pasta_pdfs}' não foi encontrada.")
        return

    dados_completos = []
    arquivos_na_pasta = [f for f in os.listdir(pasta_pdfs) if f.lower().endswith(".pdf")]

    if not arquivos_na_pasta:
        print(f"ℹ️  Nenhum arquivo PDF encontrado na pasta '{pasta_pdfs}'.")
        return

    for arquivo in arquivos_na_pasta:
        caminho_completo = os.path.join(pasta_pdfs, arquivo)
        print(f"📄 Processando arquivo: {arquivo}...")
        try:
            itens_extraidos = extrair_dados_pdf(caminho_completo)
            if itens_extraidos:
                dados_completos.extend(itens_extraidos)
                print(f"✓  Sucesso: {len(itens_extraidos)} item(ns) extraído(s).")
            else:
                print(f"⚠️  Aviso: Nenhum item correspondente ao padrão foi encontrado em {arquivo}.")
        except Exception as e:
            print(f"❌ Erro inesperado ao processar o arquivo {arquivo}: {e}")
    
    gerar_excel(dados_completos)

if __name__ == "__main__":
    print("🔄 Iniciando o processo de extração de dados dos PDFs (v2)...")
    processar_pasta()
    print("✨ Processo concluído!")
