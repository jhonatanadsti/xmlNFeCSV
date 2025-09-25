import xml.etree.ElementTree as ET
import pandas as pd
import re
from datetime import datetime

def parse_nfe_xml(xml_file_path):
    """
    Parse XML NFe e extrai dados dos produtos com informações de lote
    """

    # Parse do XML
    tree = ET.parse(xml_file_path)
    root = tree.getroot()

    # Namespaces do XML
    ns = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}

    # Dados gerais da NFe
    nfe_info = {}
    ide = root.find('.//nfe:ide', ns)
    emit = root.find('.//nfe:emit', ns)
    dest = root.find('.//nfe:dest', ns)
    total = root.find('.//nfe:total/nfe:ICMSTot', ns)
    inf_adic = root.find('.//nfe:infAdic', ns)

    if ide is not None:
        nfe_info['numero_nfe'] = ide.find('nfe:nNF', ns).text if ide.find('nfe:nNF', ns) is not None else ''
        nfe_info['serie'] = ide.find('nfe:serie', ns).text if ide.find('nfe:serie', ns) is not None else ''
        nfe_info['data_emissao'] = ide.find('nfe:dhEmi', ns).text if ide.find('nfe:dhEmi', ns) is not None else ''
        nfe_info['cfop_geral'] = ide.find('nfe:natOp', ns).text if ide.find('nfe:natOp', ns) is not None else ''

    if emit is not None:
        nfe_info['emit_cnpj'] = emit.find('nfe:CNPJ', ns).text if emit.find('nfe:CNPJ', ns) is not None else ''
        nfe_info['emit_nome'] = emit.find('nfe:xNome', ns).text if emit.find('nfe:xNome', ns) is not None else ''

    if dest is not None:
        nfe_info['dest_cnpj'] = dest.find('nfe:CNPJ', ns).text if dest.find('nfe:CNPJ', ns) is not None else ''
        nfe_info['dest_nome'] = dest.find('nfe:xNome', ns).text if dest.find('nfe:xNome', ns) is not None else ''

    if total is not None:
        nfe_info['valor_total_nfe'] = float(total.find('nfe:vNF', ns).text) if total.find('nfe:vNF', ns) is not None else 0.0
        nfe_info['icms_desonerado_total'] = float(total.find('nfe:vICMSDeson', ns).text) if total.find('nfe:vICMSDeson', ns) is not None else 0.0

    # Extrair informações de lote das informações adicionais
    lote_info = {}
    if inf_adic is not None:
        inf_cpl = inf_adic.find('nfe:infCpl', ns)
        if inf_cpl is not None:
            lote_info = parse_lote_info(inf_cpl.text)

    # Lista para armazenar dados dos produtos
    produtos_data = []

    # Primeiro, criar uma cópia dos lotes para cada produto para poder processar múltiplas linhas
    lote_info_expandido = {}
    for codigo, lotes in lote_info.items():
        lote_info_expandido[codigo] = lotes.copy()

    # Iterar sobre os produtos
    produtos = root.findall('.//nfe:det', ns)

    for produto in produtos:
        item_data = nfe_info.copy()  # Copia dados gerais

        # Dados do produto
        prod = produto.find('nfe:prod', ns)
        if prod is not None:
            item_data['item_nfe'] = produto.get('nItem', '')
            item_data['codigo_produto'] = prod.find('nfe:cProd', ns).text if prod.find('nfe:cProd', ns) is not None else ''

            item_data['descricao_produto'] = prod.find('nfe:xProd', ns).text if prod.find('nfe:xProd', ns) is not None else ''
            item_data['ncm'] = prod.find('nfe:NCM', ns).text if prod.find('nfe:NCM', ns) is not None else ''
            item_data['cfop'] = prod.find('nfe:CFOP', ns).text if prod.find('nfe:CFOP', ns) is not None else ''
            item_data['unidade_comercial'] = prod.find('nfe:uCom', ns).text if prod.find('nfe:uCom', ns) is not None else ''
            item_data['quantidade_comercial'] = float(prod.find('nfe:qCom', ns).text) if prod.find('nfe:qCom', ns) is not None else 0.0
            item_data['valor_unitario_comercial'] = float(prod.find('nfe:vUnCom', ns).text) if prod.find('nfe:vUnCom', ns) is not None else 0.0
            item_data['valor_produto'] = float(prod.find('nfe:vProd', ns).text) if prod.find('nfe:vProd', ns) is not None else 0.0
            item_data['pedido_compra'] = prod.find('nfe:xPed', ns).text if prod.find('nfe:xPed', ns) is not None else ''
            item_data['item_pedido'] = prod.find('nfe:nItemPed', ns).text if prod.find('nfe:nItemPed', ns) is not None else ''

            # Campos CEST e FCI (podem não existir)
            cest = prod.find('nfe:CEST', ns)
            item_data['cest'] = cest.text if cest is not None else ''

            fci = prod.find('nfe:nFCI', ns)
            item_data['fci'] = fci.text if fci is not None else ''

        # Dados de impostos
        imposto = produto.find('nfe:imposto', ns)
        if imposto is not None:
            # ICMS
            icms = imposto.find('.//nfe:ICMS40', ns)
            if icms is not None:
                item_data['icms_origem'] = icms.find('nfe:orig', ns).text if icms.find('nfe:orig', ns) is not None else ''
                item_data['icms_cst'] = icms.find('nfe:CST', ns).text if icms.find('nfe:CST', ns) is not None else ''
                item_data['icms_desonerado'] = float(icms.find('nfe:vICMSDeson', ns).text) if icms.find('nfe:vICMSDeson', ns) is not None else 0.0
                item_data['motivo_desoneracao'] = icms.find('nfe:motDesICMS', ns).text if icms.find('nfe:motDesICMS', ns) is not None else ''

            # IPI
            ipi = imposto.find('.//nfe:IPINT', ns)
            if ipi is not None:
                item_data['ipi_cst'] = ipi.find('nfe:CST', ns).text if ipi.find('nfe:CST', ns) is not None else ''

            # PIS
            pis = imposto.find('.//nfe:PISOutr', ns)
            if pis is not None:
                item_data['pis_cst'] = pis.find('nfe:CST', ns).text if pis.find('nfe:CST', ns) is not None else ''
                item_data['pis_base_calculo'] = float(pis.find('nfe:vBC', ns).text) if pis.find('nfe:vBC', ns) is not None else 0.0
                item_data['pis_aliquota'] = float(pis.find('nfe:pPIS', ns).text) if pis.find('nfe:pPIS', ns) is not None else 0.0
                item_data['pis_valor'] = float(pis.find('nfe:vPIS', ns).text) if pis.find('nfe:vPIS', ns) is not None else 0.0

            # COFINS
            cofins = imposto.find('.//nfe:COFINSOutr', ns)
            if cofins is not None:
                item_data['cofins_cst'] = cofins.find('nfe:CST', ns).text if cofins.find('nfe:CST', ns) is not None else ''
                item_data['cofins_base_calculo'] = float(cofins.find('nfe:vBC', ns).text) if cofins.find('nfe:vBC', ns) is not None else 0.0
                item_data['cofins_aliquota'] = float(cofins.find('nfe:pCOFINS', ns).text) if cofins.find('nfe:pCOFINS', ns) is not None else 0.0
                item_data['cofins_valor'] = float(cofins.find('nfe:vCOFINS', ns).text) if cofins.find('nfe:vCOFINS', ns) is not None else 0.0

        # Adicionar informações de lote (inicializar com valores vazios)
        item_data['infadic_produto'] = ''
        item_data['infadic_lote'] = ''
        item_data['infadic_qtd'] = ''
        item_data['infadic_unidade'] = ''

        # Buscar informações de lote
        codigo_produto = item_data.get('codigo_produto', '')

        if codigo_produto in lote_info_expandido and lote_info_expandido[codigo_produto]:
            # Pega o primeiro lote disponível para este produto
            lote_data = lote_info_expandido[codigo_produto].pop(0)
            item_data['infadic_produto'] = codigo_produto
            item_data['infadic_lote'] = lote_data['lote']
            item_data['infadic_qtd'] = lote_data['quantidade']
            item_data['infadic_unidade'] = lote_data['unidade']

        produtos_data.append(item_data)

    # Processar lotes restantes - criar linhas adicionais para produtos que têm mais lotes
    for codigo_produto, lotes_restantes in lote_info_expandido.items():
        for lote_data in lotes_restantes:
            # Encontrar um produto base para copiar os dados
            produto_base = None
            for item in produtos_data:
                if item.get('codigo_produto') == codigo_produto:
                    produto_base = item.copy()
                    break

            if produto_base:
                # Atualizar com dados do lote adicional
                produto_base['infadic_produto'] = codigo_produto
                produto_base['infadic_lote'] = lote_data['lote']
                produto_base['infadic_qtd'] = lote_data['quantidade']
                produto_base['infadic_unidade'] = lote_data['unidade']
                # Limpar dados comerciais para não duplicar valores
                produto_base['quantidade_comercial'] = 0.0
                produto_base['valor_unitario_comercial'] = 0.0
                produto_base['valor_produto'] = 0.0
                produto_base['item_nfe'] = f"{produto_base.get('item_nfe', '')}_lote_extra"

                produtos_data.append(produto_base)

    return produtos_data

def parse_lote_info(inf_cpl_text):
    """
    Parse das informações de lote do campo infCpl
    Padrão: "-100141432-LOTE: 0052246201-32SAC, 0052246203-8SAC-100141447-LOTE: ..."
    """
    lote_info = {}

    # Encontrar a seção específica de lotes que começa com "-" seguido de código
    lote_section_match = re.search(r'(-\d+-LOTE:.*)', inf_cpl_text, re.DOTALL)
    if not lote_section_match:
        return lote_info

    lote_section = lote_section_match.group(1)

    # Dividir por produtos usando regex que identifica o padrão "-CODIGO-LOTE:"
    produtos_parts = re.split(r'(?=-\d+-LOTE:)', lote_section)
    produtos_parts = [p.strip() for p in produtos_parts if p.strip()]

    for parte in produtos_parts:
        # Extrair código do produto e dados dos lotes
        match = re.match(r'-(\d+)-LOTE:\s*(.+?)(?=-\d+-LOTE:|$)', parte, re.DOTALL)
        if not match:
            continue

        codigo_produto = match.group(1)
        dados_lotes = match.group(2).strip()

        if codigo_produto not in lote_info:
            lote_info[codigo_produto] = []

        # Processar os lotes usando regex para encontrar todos os padrões LOTE-QUANTIDADEUNIDADE
        # Padrão: números-lote-quantidadewithcomma+UNIDADE
        lote_pattern = r'(\d+)-([0-9,.]+)([A-Z]{2,3})'
        lotes_encontrados = re.findall(lote_pattern, dados_lotes)
        
        for numero_lote, quantidade_str, unidade in lotes_encontrados:

            # Processar quantidade
            try:
                # Se tem vírgula seguida de exatamente 3 dígitos, é separador de milhares
                if re.match(r'^\d+,\d{3}$', quantidade_str):
                    quantidade = float(quantidade_str.replace(',', ''))
                elif re.match(r'^\d+\.\d{3}$', quantidade_str):
                    quantidade = float(quantidade_str.replace('.', ''))
                elif ',' in quantidade_str and not re.match(r'^\d+,\d{3}$', quantidade_str):
                    quantidade = float(quantidade_str.replace(',', '.'))
                elif '.' in quantidade_str and not re.match(r'^\d+\.\d{3}$', quantidade_str):
                    quantidade = float(quantidade_str.replace('.', ''))
                else:
                    # Número simples
                    quantidade = float(quantidade_str)

            except ValueError:
                quantidade = 0.0

            lote_info[codigo_produto].append({
                'lote': numero_lote,
                'quantidade': quantidade,
                'unidade': unidade
            })

    return lote_info

def main():
    """
    Função principal para executar o parser
    """
    # Caminho do arquivo XML (altere conforme necessário)
    xml_file_path = 'nfe.xml'  # Substitua pelo caminho do seu arquivo

    try:
        # Parse do XML
        produtos_data = parse_nfe_xml(xml_file_path)

        # Criar DataFrame
        df = pd.DataFrame(produtos_data)

        # Definir ordem das colunas (incluindo as novas colunas de infAdic)
        colunas_ordenadas = [
            'numero_nfe', 'serie', 'data_emissao', 'emit_cnpj', 'emit_nome',
            'dest_cnpj', 'dest_nome', 'valor_total_nfe', 'icms_desonerado_total',
            'item_nfe', 'codigo_produto', 'descricao_produto', 'ncm', 'cest', 'fci',
            'cfop', 'unidade_comercial', 'quantidade_comercial', 'valor_unitario_comercial',
            'valor_produto', 'pedido_compra', 'item_pedido',
            'infadic_produto', 'infadic_lote', 'infadic_qtd', 'infadic_unidade',
            'icms_origem', 'icms_cst', 'icms_desonerado', 'motivo_desoneracao',
            'ipi_cst', 'pis_cst', 'pis_base_calculo', 'pis_aliquota', 'pis_valor',
            'cofins_cst', 'cofins_base_calculo', 'cofins_aliquota', 'cofins_valor'
        ]

        # Reordenar colunas (apenas as que existem)
        colunas_existentes = [col for col in colunas_ordenadas if col in df.columns]
        df = df[colunas_existentes]

        # Exibir resultado
        print("Dados extraídos da NFe:")
        print(f"Total de itens: {len(df)}")
        print("\nPrimeiras linhas:")
        print(df.head())

        # Verificar se as colunas de infAdic foram preenchidas
        print("\nVerificação das colunas infAdic:")
        print(f"Produtos com lote preenchido: {df['infadic_lote'].notna().sum()}")
        print(f"Produtos com lote não vazio: {(df['infadic_lote'] != '').sum()}")

        print("\nAmostras das colunas infAdic:")
        infadic_cols = ['infadic_produto', 'infadic_lote', 'infadic_qtd', 'infadic_unidade']
        sample_data = df[infadic_cols].head(10)
        for index, row in sample_data.iterrows():
            if any(str(row[col]) != '' for col in infadic_cols):
                print(f"Item {index + 1}: Produto={row['infadic_produto']}, Lote={row['infadic_lote']}, Qtd={row['infadic_qtd']}, Unid={row['infadic_unidade']}")

        # Salvar em arquivo Excel
        output_file = 'nfe_dados_completos.xlsx'
        df.to_excel(output_file, index=False, engine='openpyxl')
        print(f"\nDados salvos em: {output_file}")

        # Salvar em CSV também
        output_csv = 'nfe_dados_completos.csv'
        df.to_csv(output_csv, index=False, encoding='utf-8-sig', sep=';')
        print(f"Dados salvos em: {output_csv}")

        # Exibir estatísticas
        print("\nResumo por produto:")
        resumo = df.groupby(['codigo_produto', 'descricao_produto']).agg({
            'quantidade_comercial': 'sum',
            'valor_produto': 'sum',
            'item_nfe': 'count'
        }).round(2)
        resumo.columns = ['Quantidade Total', 'Valor Total', 'Qtd Itens']
        print(resumo)

        return df

    except FileNotFoundError:
        print(f"Arquivo {xml_file_path} não encontrado.")
        print("Por favor, certifique-se de que o arquivo XML está no mesmo diretório do script.")
        return None
    except ET.ParseError as e:
        print(f"Erro ao fazer parse do XML: {e}")
        return None
    except Exception as e:
        print(f"Erro inesperado: {e}")
        return None

# Função para testar com string XML diretamente
def test_with_xml_string():
    """
    Teste com o XML fornecido como string
    """
    xml_content = '''<nfeProc xmlns="http://www.portalfiscal.inf.br/nfe" versao="4.00">
    <!-- Cole aqui o XML completo da NFe -->
    </nfeProc>'''

    # Salvar XML em arquivo temporário para teste
    with open('temp_nfe.xml', 'w', encoding='utf-8') as f:
        f.write(xml_content)

    # Parse do arquivo temporário
    return parse_nfe_xml('temp_nfe.xml')

# Função para testar apenas o parsing de lotes
def test_lote_parsing():
    """
    Função para testar apenas a extração de lotes com o exemplo fornecido
    """
    sample_text = """S/PED:4517516666 100141432: Produto produzido a partir de milho transgenico - Por favor, caso venha a receber e-mails supostamente enviados pela Cargill solicitando a substituicao de boletos bancarios, acione imediatamente o seu contato da area comercial da Empresa. 100141432: Produto produzido a partir de milho transgenico 100141447: Produto produzido a partir de milho transgenico 100141447: Produto produzido a partir de milho transgenico 100141447: Produto produzido a partir de milho transgenico 100141493: Produto produzido a partir de milho transgenico 100141493: Produto produzido a partir de milho transgenico Valor do ICMS desonerado = R$ 34989.22 - 000001 .Nao incidencia de ICMS conforme Artigo 6o, inciso I do RICMS/AL. -100141432-LOTE: 0052246201-32SAC, 0052246203-8SAC-100141447-LOTE: 0052132134-500SAC, 0052132136-1,300SAC, 0052132139-900SAC-100141493-LOTE: 0051890670-10SAC, 0051924544-10SAC-100142227-LOTE: 0051612259-16.000TAM-100142397-LOTE: 0051428952-700SAC"""

    print("TESTE DE PARSING DE LOTES")
    print("=" * 60)
    result = parse_lote_info(sample_text)
    print("Resultado:", result)
    return result

if __name__ == "__main__":
    # Descomente a linha abaixo para testar apenas o parsing de lotes
    # test_lote_parsing()

    # Execução normal
    df_result = main()
