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
    emit = inf_nfe.find('.//nfe:emit', ns)  # Fornecedor (emitente)
    dest = inf_nfe.find('.//nfe:dest', ns)
    total = inf_nfe.find('.//nfe:total/nfe:ICMSTot', ns)

    # Extract supplier (fornecedor) information
    forn_razao = emit.find('nfe:xNome', ns).text if emit.find('nfe:xNome', ns) is not None else ""
    forn_cnpj = emit.find('nfe:CNPJ', ns).text if emit.find('nfe:CNPJ', ns) is not None else ""
    forn_ie = emit.find('nfe:IE', ns).text if emit.find('nfe:IE', ns) is not None else ""
    forn_endereco = emit.find('.//nfe:enderEmit/nfe:xLgr', ns).text if emit.find('.//nfe:enderEmit/nfe:xLgr', ns) is not None else ""
    forn_bairro = emit.find('.//nfe:enderEmit/nfe:xBairro', ns).text if emit.find('.//nfe:enderEmit/nfe:xBairro', ns) is not None else ""
    forn_cidade = emit.find('.//nfe:enderEmit/nfe:xMun', ns).text if emit.find('.//nfe:enderEmit/nfe:xMun', ns) is not None else ""
    forn_uf = emit.find('.//nfe:enderEmit/nfe:UF', ns).text if emit.find('.//nfe:enderEmit/nfe:UF', ns) is not None else ""
    forn_cep = emit.find('.//nfe:enderEmit/nfe:CEP', ns).text if emit.find('.//nfe:enderEmit/nfe:CEP', ns) is not None else ""

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
        'nf_dt_emissao': "",  # Ensure empty
        'nf_hora': "",        # Ensure empty
        'nf_dt_entrada': "",  # Ensure empty
        'nf_horaentrada': "", # Ensure empty
        'nf_cfop': "",        # Will be filled from items
        'nf_obs': inf_nfe.find('.//nfe:infAdic/nfe:infCpl', ns).text if inf_nfe.find('.//nfe:infAdic/nfe:infCpl', ns) is not None else "",
        'nf_base_icms': total.find('nfe:vBC', ns).text if total.find('nfe:vBC', ns) is not None else "0",
        'nf_valor_icms': total.find('nfe:vICMS', ns).text if total.find('nfe:vICMS', ns) is not None else "0",
        'nf_valor_total': total.find('nfe:vNF', ns).text if total.find('nfe:vNF', ns) is not None else "0",
        'nf_valor_total_prod': total.find('nfe:vProd', ns).text if total.find('nfe:vProd', ns) is not None else "0",
        
        # Client information
        'cli_razao': "",      # Ensure empty
        'cli_cnpj': "",       # Ensure empty
        'cli_ie': "",         # Ensure empty
        'cli_endereco': "",   # Ensure empty
        'cli_bairro': "",     # Ensure empty
        'cli_cidade': "",     # Ensure empty
        'cli_uf': "",         # Ensure empty
        'cli_cep': "",        # Ensure empty

        # Supplier (fornecedor) information
        'forn_razao': forn_razao,
        'forn_cnpj': forn_cnpj,
        'forn_ie': forn_ie,
        'forn_endereco': forn_endereco,
        'forn_bairro': forn_bairro,
        'forn_cidade': forn_cidade,
        'forn_uf': forn_uf,
        'forn_cep': forn_cep,
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
            'item_serial': "",     # Ensure empty
            'item_modelo': "",     # Ensure empty
            'item_valor_unit': prod.find('nfe:vUnCom', ns).text if prod.find('nfe:vUnCom', ns) is not None else "0",
            'item_valor_total': prod.find('nfe:vProd', ns).text if prod.find('nfe:vProd', ns) is not None else "0",
            'item_valor_icms': "", # Ensure empty
            'item_valor_ipi': "",  # Ensure empty
            'item_aliq_icms': "",  # Ensure empty
            'item_aliq_ipi': "",   # Ensure empty
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
    
    # Select only the specified columns
    columns = [
        'nf_numnota',      # Número da Nota Fiscal
        'nf_serie',        # Série da Nota Fiscal
        'nf_dt_emissao',   # Data de Emissão
        'nf_hora',         # Hora de Emissão
        'nf_dt_entrada',   # Data de Entrada
        'nf_horaentrada',  # Hora de Entrada
        'nf_cfop',         # CFOP
        'nf_obs',          # Observações
        'nf_base_icms',    # Base ICMS
        'nf_valor_icms',   # Valor ICMS
        'nf_valor_total',  # Valor Total
        'nf_valor_total_prod', # Valor Total dos Produtos
        'cli_razao',       # Razão Social do Cliente
        'cli_cnpj',        # CNPJ do Cliente
        'cli_ie',          # Inscrição Estadual do Cliente
        'cli_endereco',    # Endereço do Cliente
        'cli_bairro',      # Bairro do Cliente
        'cli_cidade',      # Cidade do Cliente
        'cli_uf',          # UF do Cliente
        'cli_cep',         # CEP do Cliente
        'forn_razao',      # Razão Social do Fornecedor
        'forn_cnpj',       # CNPJ do Fornecedor
        'forn_ie',         # Inscrição Estadual do Fornecedor
        'forn_endereco',   # Endereço do Fornecedor
        'forn_bairro',     # Bairro do Fornecedor
        'forn_cidade',     # Cidade do Fornecedor
        'forn_uf',         # UF do Fornecedor
        'forn_cep',        # CEP do Fornecedor
        'item_codigo',     # Código do Item
        'item_descricao',  # Descrição do Item
        'item_ncm',        # NCM do Item
        'item_un',         # Unidade do Item
        'item_qtde',       # Quantidade do Item
        'item_lote',       # Lote do Item
        'item_serial',     # Serial do Item
        'item_modelo',     # Modelo do Item
        'item_valor_unit', # Valor Unitário do Item
        'item_valor_total',# Valor Total do Item
        'item_valor_icms', # Valor ICMS do Item
        'item_valor_ipi',  # Valor IPI do Item
        'item_aliq_icms',  # Alíquota ICMS do Item
        'item_aliq_ipi'    # Alíquota IPI do Item
    ]
    
    # Ensure all columns exist (with empty values if needed)
    for col in columns:
        if col not in df.columns:
            df[col] = ""
    
    # Reorder columns to ensure they appear in the CSV in the correct order
    df = df[columns]
    
    # Ensure specific columns are treated as text
    text_columns = [
        'forn_razao', 
        'forn_endereco', 'forn_bairro', 'forn_cidade', 
        'item_descricao', 'cli_razao', 'cli_cnpj', 'cli_ie', 
        'cli_endereco', 'cli_bairro', 'cli_cidade'
    ]
    for col in text_columns:
        if col in df.columns:
            df[col] = df[col].astype(str)
    
    # Remove explicit formatting for `forn_cnpj` and `forn_ie`
    # These columns will retain their original format from the XML
    
    # Format numeric columns to replace '.' with ',' for decimal separator
    numeric_columns = [
        'nf_base_icms', 'nf_valor_icms', 'nf_valor_total', 
        'nf_valor_total_prod', 'item_valor_unit', 'item_valor_total', 
        'item_qtde', 'item_valor_icms', 'item_valor_ipi', 
        'item_aliq_icms', 'item_aliq_ipi'
    ]
    for col in numeric_columns:
        if col in df.columns:
            # Replace '.' with ',' without changing the structure of the number
            df[col] = df[col].astype(str).str.replace('.', ',', regex=False)
    
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
    # Set page configuration with logo and new title
    st.set_page_config(
        page_title="Conversor XML-CSV Solution",
        page_icon="Logo Solution.png",
        layout="wide"
    )
    
    # Add logo to the sidebar
    st.sidebar.image("Logo Solution.png", use_container_width=True)
    
    # Update the title
    st.title("Conversor XML-CSV Solution")
    st.write("Faça upload de arquivos XML de Nota Fiscal Eletrônica (NFe) para convertê-los para CSV em formato tabular.")
    
    # Allow multiple file uploads
    uploaded_files = st.file_uploader("Escolha os arquivos XML", type="xml", accept_multiple_files=True)
    
    if uploaded_files:
        try:
            all_data = []  # List to store data from all XML files
            
            # Load laborlog.xlsx for comparison
            laborlog_path = "laborlog.xlsx"  # Ensure this file exists in the same directory
            laborlog_df = pd.read_excel(laborlog_path)
            
            # Iterate over uploaded files
            for uploaded_file in uploaded_files:
                # Read XML content
                xml_content = uploaded_file.read().decode('utf-8')
                
                # Clean XML content (remove BOM if present)
                xml_content = re.sub(r'^\xef\xbb\xbf', '', xml_content)
                
                with st.spinner(f"Processando arquivo {uploaded_file.name}..."):
                    # Parse XML and generate data
                    parsed_data = parse_nfe_xml(xml_content)
                    
                    if parsed_data:
                        all_data.extend(parsed_data)
                    else:
                        st.error(f"Falha ao analisar o arquivo {uploaded_file.name}. Verifique se é um arquivo XML de NFe válido.")
            
            if all_data:
                # Generate DataFrame from all parsed data
                df = generate_csv(all_data)
                
                # Perform "PROCV" operation
                if 'EAN' in laborlog_df.columns and 'CÓD. LABORLOG' in laborlog_df.columns:
                    ean_to_cod = laborlog_df.set_index('EAN')['CÓD. LABORLOG'].to_dict()
                    # Map 'forn_ie' to 'item_codigo' and replace missing values with "ERRO"
                    df['item_codigo'] = df['forn_ie'].map(ean_to_cod).fillna("ERRO")
                
                # Display sample of data in tabular format
                st.subheader("Visualização dos Dados Convertidos")
                st.dataframe(df.head(10))
                
                # Display download options
                st.subheader("Download")
                csv_data = df.to_csv(index=False, sep=';', encoding='utf-8-sig')
                st.download_button(
                    label="Download CSV File",
                    data=csv_data,
                    file_name="nfe_data.csv",
                    mime="text/csv"
                )
            else:
                st.error("Nenhum dado válido foi extraído dos arquivos XML.")
        
        except Exception as e:
            st.error(f"Erro ao processar os arquivos: {str(e)}")
            st.write("Certifique-se de que está enviando arquivos XML de NFe válidos e que o arquivo laborlog.xlsx está presente.")

if __name__ == "__main__":
    main()
