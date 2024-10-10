import xml.etree.ElementTree as ET
import pandas as pd
import os
import glob

def extrair_produtos_nfe(xml_file):
    ns = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}
    tree = ET.parse(xml_file)
    root = tree.getroot()

    produtos = []
    for det in root.findall('.//nfe:det', ns):
        prod = det.find('nfe:prod', ns)
        cProd = prod.find('nfe:cProd', ns).text
        xProd = prod.find('nfe:xProd', ns).text
        qCom = prod.find('nfe:qCom', ns).text
        vUnCom = prod.find('nfe:vUnCom', ns).text
        vProd = prod.find('nfe:vProd', ns).text
        NCM = prod.find('nfe:NCM', ns).text
        CFOP = prod.find('nfe:CFOP', ns).text

        produtos.append({
            'Código do Produto': cProd,
            'Descrição': xProd,
            'Quantidade': qCom,
            'CFOP': CFOP,
            'NCM': NCM,
            'Valor Unitário': vUnCom,
            'Valor Total': vProd,
            'Arquivo XML': os.path.basename(xml_file)  # Identificar de qual arquivo XML veio o produto
        })

    return produtos

def salvar_produtos_em_unico_excel(pasta_xml, output_file):
    todos_produtos = []

    # Buscar todos os arquivos XML na pasta
    xml_files = glob.glob(os.path.join(pasta_xml, "*.xml"))

    for xml_file in xml_files:
        produtos = extrair_produtos_nfe(xml_file)
        todos_produtos.extend(produtos)

    # Salva todos os produtos em um arquivo Excel
    df = pd.DataFrame(todos_produtos)
    df.to_excel(output_file, index=False)

# Exemplo de uso:
pasta_xml = '/content/XML'  # Substituir pelo caminho da pasta
output_file = 'produtos_todos_nfes.xlsx'

# Salvar todos os produtos de vários arquivos XML em um único Excel
salvar_produtos_em_unico_excel(pasta_xml, output_file)

print(f"Produtos extraídos de todos os arquivos XML foram salvos em: {output_file}")
