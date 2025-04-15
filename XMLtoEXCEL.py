import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
import base64
from datetime import datetime
import io
from io import BytesIO
import re

def parse_nfe_xml(xml_content):
    """
    Parse NFe XML content and extract relevant data.
    """
    # Define namespace
    ns = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}
    
    # Parse XML content
    try:
        root = ET.fromstring(xml_content)
    except ET.ParseError:
        st.error("Invalid XML format")
        return None
    
    # Find NFe element
    if 'nfeProc' in root.tag:
        nfe = root.find('.//nfe:NFe', ns)
    else:
        nfe = root if 'NFe' in root.tag else None
    
    if nfe is None:
        st.error("NFe not found in XML")
        return None
    
    # Extract infNFe data
    inf_nfe = nfe.find('.//nfe:infNFe', ns)
    if inf_nfe is None:
        st.error("infNFe not found in XML")
        return None
    
    # Extract general invoice information
    ide = inf_nfe.find('.//nfe:ide', ns)
    emit = inf_nfe.find('.//nfe:emit', ns)
    dest = inf_nfe.find('.//nfe:dest', ns)
    total = inf_nfe.find('.//nfe:total/nfe:ICMSTot', ns)
    
    # Get emission date and time
    dhEmi = ide.find('nfe:dhEmi', ns).text if ide.find('nfe:dhEmi', ns) is not None else ""
    dt_emissao = ""
    hora_emissao = ""
    if dhEmi:
        try:
            dt_obj = datetime.fromisoformat(dhEmi.replace('Z', '+00:00'))
            dt_emissao = dt_obj.strftime('%Y-%m-%d')
            hora_emissao = dt_obj.strftime('%H:%M:%S')
        except ValueError:
            dt_emissao = dhEmi.split('T')[0] if 'T' in dhEmi else ""
            hora_emissao = dhEmi.split('T')[1].split('-')[0] if 'T' in dhEmi else ""
    
    # Create general invoice data
    invoice_data = {
        'nf_numnota': ide.find('nfe:nNF', ns).text if ide.find('nfe:nNF', ns) is not None else "",
        'nf_serie': ide.find('nfe:serie', ns).text if ide.find('nfe:serie', ns) is not None else "",
        'nf_dt_emissao': dt_emissao,
        'nf_hora': hora_emissao,
        'nf_dt_entrada': dt_emissao,  # Using emission date as entry date (not always provided in XML)
        'nf_horaentrada': hora_emissao,  # Using emission time as entry time (not always provided in XML)
        'nf_cfop': "",  # Will be filled from items
        'nf_obs': inf_nfe.find('.//nfe:infAdic/nfe:infCpl', ns).text if inf_nfe.find('.//nfe:infAdic/nfe:infCpl', ns) is not None else "",
        'nf_base_icms': total.find('nfe:vBC', ns).text if total.find('nfe:vBC', ns) is not None else "0",
        'nf_valor_icms': total.find('nfe:vICMS', ns).text if total.find('nfe:vICMS', ns) is not None else "0",
        'nf_valor_total': total.find('nfe:vNF', ns).text if total.find('nfe:vNF', ns) is not None else "0",
        'nf_valor_total_prod': total.find('nfe:vProd', ns).text if total.find('nfe:vProd', ns) is not None else "0",
        
        # Client information
        'cli_razao': dest.find('nfe:xNome', ns).text if dest.find('nfe:xNome', ns) is not None else "",
        'cli_cnpj': dest.find('nfe:CNPJ', ns).text if dest.find('nfe:CNPJ', ns) is not None else (dest.find('nfe:CPF', ns).text if dest.find('nfe:CPF', ns) is not None else ""),
        'cli_ie': dest.find('nfe:IE', ns).text if dest.find('nfe:IE', ns) is not None else "",
        'cli_endereco': f"{dest.find('.//nfe:enderDest/nfe:xLgr', ns).text} {dest.find('.//nfe:enderDest/nfe:nro', ns).text}" if dest.find('.//nfe:enderDest/nfe:xLgr', ns) is not None else "",
        'cli_bairro': dest.find('.//nfe:enderDest/nfe:xBairro', ns).text if dest.find('.//nfe:enderDest/nfe:xBairro', ns) is not None else "",
        'cli_cidade': dest.find('.//nfe:enderDest/nfe:xMun', ns).text if dest.find('.//nfe:enderDest/nfe:xMun', ns) is not None else "",
        'cli_uf': dest.find('.//nfe:enderDest/nfe:UF', ns).text if dest.find('.//nfe:enderDest/nfe:UF', ns) is not None else "",
        'cli_cep': dest.find('.//nfe:enderDest/nfe:CEP', ns).text if dest.find('.//nfe:enderDest/nfe:CEP', ns) is not None else "",
        
        # Supplier information
        'forn_razao': emit.find('nfe:xNome', ns).text if emit.find('nfe:xNome', ns) is not None else "",
        'forn_cnpj': emit.find('nfe:CNPJ', ns).text if emit.find('nfe:CNPJ', ns) is not None else (emit.find('nfe:CPF', ns).text if emit.find('nfe:CPF', ns) is not None else ""),
        'forn_ie': emit.find('nfe:IE', ns).text if emit.find('nfe:IE', ns) is not None else "",
        'forn_endereco': f"{emit.find('.//nfe:enderEmit/nfe:xLgr', ns).text} {emit.find('.//nfe:enderEmit/nfe:nro', ns).text}" if emit.find('.//nfe:enderEmit/nfe:xLgr', ns) is not None else "",
        'forn_bairro': emit.find('.//nfe:enderEmit/nfe:xBairro', ns).text if emit.find('.//nfe:enderEmit/nfe:xBairro', ns) is not None else "",
        'forn_cidade': emit.find('.//nfe:enderEmit/nfe:xMun', ns).text if emit.find('.//nfe:enderEmit/nfe:xMun', ns) is not None else "",
        'forn_uf': emit.find('.//nfe:enderEmit/nfe:UF', ns).text if emit.find('.//nfe:enderEmit/nfe:UF', ns) is not None else "",
        'forn_cep': emit.find('.//nfe:enderEmit/nfe:CEP', ns).text if emit.find('.//nfe:enderEmit/nfe:CEP', ns) is not None else "",
    }
    
    # Extract items
    items = []
    for det in inf_nfe.findall('.//nfe:det', ns):
        prod = det.find('nfe:prod', ns)
        imposto = det.find('nfe:imposto', ns)
        
        # Get ICMS data
        icms_tag = imposto.find('.//nfe:ICMS', ns)
        icms_values = {}
        if icms_tag is not None:
            # Try different ICMS types
            for icms_type in ['ICMS00', 'ICMS10', 'ICMS20', 'ICMS30', 'ICMS40', 'ICMS51', 'ICMS60', 'ICMS70', 'ICMS90']:
                icms_specific = icms_tag.find(f'nfe:{icms_type}', ns)
                if icms_specific is not None:
                    icms_values['vICMS'] = icms_specific.find('nfe:vICMS', ns).text if icms_specific.find('nfe:vICMS', ns) is not None else "0"
                    icms_values['pICMS'] = icms_specific.find('nfe:pICMS', ns).text if icms_specific.find('nfe:pICMS', ns) is not None else "0"
                    break
        
        # Get IPI data
        ipi_tag = imposto.find('.//nfe:IPI', ns)
        ipi_values = {"vIPI": "0", "pIPI": "0"}
        if ipi_tag is not None:
            # Try different IPI types
            for ipi_type in ['IPITrib']:
                ipi_specific = ipi_tag.find(f'nfe:{ipi_type}', ns)
                if ipi_specific is not None:
                    ipi_values['vIPI'] = ipi_specific.find('nfe:vIPI', ns).text if ipi_specific.find('nfe:vIPI', ns) is not None else "0"
                    ipi_values['pIPI'] = ipi_specific.find('nfe:pIPI', ns).text if ipi_specific.find('nfe:pIPI', ns) is not None else "0"
                    break
        
        # Get item lot information
        rastro = prod.find('nfe:rastro', ns)
        lote = ""
        if rastro is not None:
            lote = rastro.find('nfe:nLote', ns).text if rastro.find('nfe:nLote', ns) is not None else ""
        
        # CFOP for invoice
        cfop = prod.find('nfe:CFOP', ns).text if prod.find('nfe:CFOP', ns) is not None else ""
        if cfop and not invoice_data['nf_cfop']:
            invoice_data['nf_cfop'] = cfop
        
        # Extract product data
        item = {
            'item_codigo': prod.find('nfe:cProd', ns).text if prod.find('nfe:cProd', ns) is not None else "",
            'item_descricao': prod.find('nfe:xProd', ns).text if prod.find('nfe:xProd', ns) is not None else "",
            'item_ncm': prod.find('nfe:NCM', ns).text if prod.find('nfe:NCM', ns) is not None else "",
            'item_un': prod.find('nfe:uCom', ns).text if prod.find('nfe:uCom', ns) is not None else "",
            'item_qtde': prod.find('nfe:qCom', ns).text if prod.find('nfe:qCom', ns) is not None else "0",
            'item_lote': lote,
            'item_serial': "",  # Not usually specified in NFe
            'item_modelo': "",  # Not usually specified in NFe
            'item_valor_unit': prod.find('nfe:vUnCom', ns).text if prod.find('nfe:vUnCom', ns) is not None else "0",
            'item_valor_total': prod.find('nfe:vProd', ns).text if prod.find('nfe:vProd', ns) is not None else "0",
            'item_valor_icms': icms_values.get('vICMS', "0"),
            'item_valor_ipi': ipi_values.get('vIPI', "0"),
            'item_aliq_icms': icms_values.get('pICMS', "0"),
            'item_aliq_ipi': ipi_values.get('pIPI', "0"),
        }
        
        items.append({**invoice_data, **item})
    
    return items

def generate_csv(data):
    """
    Convert parsed data to CSV format with specific column order.
    """
    if not data:
        return None
    
    df = pd.DataFrame(data)
    
    # Order columns as requested (Coluna A, B, C, etc.)
    columns = [
        'nf_numnota',      # Coluna A
        'nf_serie',        # Coluna B
        'nf_dt_emissao',   # Coluna C
        'nf_hora',         # Coluna D
        'nf_dt_entrada',   # Coluna E
        'nf_horaentrada',  # Coluna F
        'nf_cfop',         # Coluna G
        'nf_obs',          # Coluna H
        'nf_base_icms',    # Coluna I
        'nf_valor_icms',   # Coluna J
        'nf_valor_total',  # Coluna K
        'nf_valor_total_prod', # Coluna L
        'cli_razao',       # Coluna M
        'cli_cnpj',        # Coluna N
        'cli_ie',          # Coluna O
        'cli_endereco',    # Coluna P
        'cli_bairro',      # Coluna Q
        'cli_cidade',      # Coluna R
        'cli_uf',          # Coluna S
        'cli_cep',         # Coluna T
        'forn_razao',      # Coluna U
        'forn_cnpj',       # Coluna V
        'forn_ie',         # Coluna W
        'forn_endereco',   # Coluna X
        'forn_bairro',     # Coluna Y
        'forn_cidade',     # Coluna Z
        'forn_uf',         # Coluna AA
        'forn_cep',        # Coluna AB
        'item_codigo',     # Coluna AC
        'item_descricao',  # Coluna AD
        'item_ncm',        # Coluna AE
        'item_un',         # Coluna AF
        'item_qtde',       # Coluna AG
        'item_lote',       # Coluna AH
        'item_serial',     # Coluna AI
        'item_modelo',     # Coluna AJ
        'item_valor_unit', # Coluna AK
        'item_valor_total', # Coluna AL
        'item_valor_icms', # Coluna AM
        'item_valor_ipi',  # Coluna AN
        'item_aliq_icms',  # Coluna AO
        'item_aliq_ipi'    # Coluna AP
    ]
    
    # Ensure all columns exist (with empty values if needed)
    for col in columns:
        if col not in df.columns:
            df[col] = ""
    
    # Reorder columns to ensure they appear in the CSV in the correct order
    df = df[columns]
    
    return df

def create_download_link(df, filename="nfe_data.csv"):
    """
    Create a download link for the dataframe.
    """
    # Use index=False to prevent extra column with row numbers
    # sep=';' to ensure proper separation for Brazilian CSV format
    csv = df.to_csv(index=False, sep=';', encoding='utf-8-sig')
    b64 = base64.b64encode(csv.encode()).decode()
    href = f'<a href="data:file/csv;base64,{b64}" download="{filename}">Download CSV File</a>'
    return href

def show_column_mapping(columns):
    """
    Display the mapping of columns to Excel-style column letters.
    """
    excel_cols = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']
    # For columns after Z: AA, AB, AC, etc.
    for i in range(26, 100):
        letter = excel_cols[i // 26 - 1] + excel_cols[i % 26]
        excel_cols.append(letter)
    
    mapping = []
    for i, col in enumerate(columns):
        mapping.append(f"Coluna {excel_cols[i]}: {col}")
    
    return mapping

def main():
    st.set_page_config(page_title="NFe XML para CSV - Formato Tabular", layout="wide")
    
    st.title("Conversor de NFe XML para CSV")
    st.write("Faça upload de um arquivo XML de Nota Fiscal Eletrônica (NFe) para convertê-lo para CSV em formato tabular.")
    
    uploaded_file = st.file_uploader("Escolha um arquivo XML", type="xml")
    
    # Define column order
    columns = [
        'nf_numnota', 'nf_serie', 'nf_dt_emissao', 'nf_hora', 
        'nf_dt_entrada', 'nf_horaentrada', 'nf_cfop', 'nf_obs', 
        'nf_base_icms', 'nf_valor_icms', 'nf_valor_total', 'nf_valor_total_prod',
        'cli_razao', 'cli_cnpj', 'cli_ie', 'cli_endereco', 'cli_bairro', 
        'cli_cidade', 'cli_uf', 'cli_cep',
        'forn_razao', 'forn_cnpj', 'forn_ie', 'forn_endereco', 'forn_bairro', 
        'forn_cidade', 'forn_uf', 'forn_cep',
        'item_codigo', 'item_descricao', 'item_ncm', 'item_un', 'item_qtde', 
        'item_lote', 'item_serial', 'item_modelo', 'item_valor_unit', 
        'item_valor_total', 'item_valor_icms', 'item_valor_ipi', 
        'item_aliq_icms', 'item_aliq_ipi'
    ]
    
    # Show column mapping in expandable section
    with st.expander("Ver mapeamento de colunas"):
        st.write("O arquivo CSV gerado terá as seguintes colunas na ordem (equivalentes a colunas do Excel):")
        mapping = show_column_mapping(columns)
        for i, map_item in enumerate(mapping):
            st.write(map_item)
    
    if uploaded_file is not None:
        try:
            # Read XML content
            xml_content = uploaded_file.read().decode('utf-8')
            
            # Clean XML content (remove BOM if present)
            xml_content = re.sub(r'^\xef\xbb\xbf', '', xml_content)
            
            with st.spinner("Processando arquivo XML..."):
                # Parse XML and generate dataframe
                parsed_data = parse_nfe_xml(xml_content)
                
                if parsed_data:
                    df = generate_csv(parsed_data)
                    
                    # Display sample of data in tabular format
                    st.subheader("Visualização dos Dados Convertidos")
                    st.dataframe(df.head(10))
                    
                    # Display download options
                    st.subheader("Download")
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        st.markdown(create_download_link(df, "nfe_data.csv"), unsafe_allow_html=True)
                    
                    with col2:
                        # Option to generate Excel file instead
                        buffer = io.BytesIO()
                        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                            df.to_excel(writer, sheet_name='NFe Data', index=False)
                        buffer.seek(0)
                        
                        st.download_button(
                            label="Download Excel File",
                            data=buffer,
                            file_name="nfe_data.xlsx",
                            mime="application/vnd.ms-excel"
                        )
                    
                    # Display additional statistics
                    st.subheader("Resumo da Nota Fiscal")
                    col1, col2, col3 = st.columns(3)
                    
                    with col1:
                        st.metric("Número da NF", df['nf_numnota'].iloc[0])
                        st.metric("Data de Emissão", df['nf_dt_emissao'].iloc[0])
                    
                    with col2:
                        st.metric("Valor Total", f"R$ {float(df['nf_valor_total'].iloc[0]):,.2f}")
                        st.metric("Valor Total Produtos", f"R$ {float(df['nf_valor_total_prod'].iloc[0]):,.2f}")
                    
                    with col3:
                        st.metric("Valor ICMS", f"R$ {float(df['nf_valor_icms'].iloc[0]):,.2f}")
                        st.metric("Número de Itens", len(df))
                    
                    # Display additional information
                    with st.expander("Detalhes da Nota Fiscal"):
                        st.write(f"**Fornecedor:** {df['forn_razao'].iloc[0]} (CNPJ: {df['forn_cnpj'].iloc[0]})")
                        st.write(f"**Cliente:** {df['cli_razao'].iloc[0]} (CNPJ: {df['cli_cnpj'].iloc[0]})")
                        st.write(f"**CFOP:** {df['nf_cfop'].iloc[0]}")
                        if df['nf_obs'].iloc[0]:
                            st.write(f"**Observações:** {df['nf_obs'].iloc[0]}")
                else:
                    st.error("Falha ao analisar os dados XML. Verifique se é um arquivo XML de NFe válido.")
        
        except Exception as e:
            st.error(f"Erro ao processar o arquivo: {str(e)}")
            st.write("Certifique-se de que está enviando um arquivo XML de NFe válido.")

if __name__ == "__main__":
    main()