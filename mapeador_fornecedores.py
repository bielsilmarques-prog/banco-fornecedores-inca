import requests
import pandas as pd
import time
import re

def consultar_dados_cnpj(cnpj):
    cnpj_limpo = re.sub(r'\D', '', str(cnpj or ''))
    if len(cnpj_limpo) != 14:
        return {}
    
    url = f"https://brasilapi.com.br/api/cnpj/v1/{cnpj_limpo}"
    try:
        res = requests.get(url, timeout=10)
        if res.status_code == 200:
            dados = res.json()
            ddd = dados.get("ddd_telefone_1", "")
            tel = dados.get("telefone_1", "")
            telefone_formatado = f"({ddd}) {tel}" if ddd and tel else tel
            
            return {
                "E-mail Comercial": dados.get("email", ""),
                "Telefone Comercial": telefone_formatado,
                "Nome Fantasia": dados.get("nome_fantasia", ""),
                "UF Fornecedor": dados.get("uf", ""),
                "Município Fornecedor": dados.get("municipio", ""),
                "Situação Cadastral": dados.get("descricao_situacao_cadastral", "")
            }
    except Exception as e:
        print(f"   [!] Erro ao buscar CNPJ {cnpj_limpo}: {e}")
    
    return {}


def buscar_todos_medicamentos_brasil(data_inicio, data_fim, paginas_max=20):
    url_pncp = "https://pncp.gov.br/api/consulta/v1/contratacoes/publicacao"
    
    # Modalidades mais comuns de compras públicas de medicamentos e saúde
    modalidades = ["6", "8", "9", "10", "14"]
    
    # Termos ultragenericos para capturar QUALQUER tipo de medicamento, farmacia e insumo
    termos_medicamentos = [
        "medicament", "medicacao", "medicação", "fármac", "farmac", "droga", 
        "remedio", "remédio", "solucao injetavel", "solução injetável", 
        "comprimido", "ampola", "frasco", "vacina", "insumo farmaceutico", 
        "insumo farmacêutico", "oncolog", "quimioterap", "antibiotico", 
        "imunoglobulina", "soro", "câncer", "cancer", "inca", "saude", "saúde"
    ]
    
    resultados = []
    print("=" * 80)
    print("INICIANDO MAPEAMENTO NACIONAL DE TODOS OS FORNECEDORES DE MEDICAMENTOS (PNCP)")
    print("=" * 80)
    
    for modalidade in modalidades:
        print(f"\n[+] Varrendo Modalidade de Contratação {modalidade}...")
        for pagina in range(1, paginas_max + 1):
            params = {
                "dataInicial": data_inicio,
                "dataFinal": data_fim,
                "codigoModalidadeContratacao": modalidade,
                "pagina": pagina,
                "tamanhoPagina": 50
            }
            
            try:
                res = requests.get(url_pncp, params=params, timeout=15)
                
                if res.status_code in [204, 404]:
                    break
                if res.status_code != 200:
                    print(f"   [-] Status {res.status_code} na página {pagina}")
                    break
                
                dados_json = res.json()
                itens = dados_json.get('data', []) if isinstance(dados_json, dict) else []
                if not itens:
                    break
                
                for item in itens:
                    orgao = item.get('orgaoEntidade', {}).get('razaoSocial', '')
                    unidade = item.get('unidadeOrgao', {}).get('nomeUnidade', '')
                    objeto = item.get('objetoContratacao', '')
                    texto_analise = f"{orgao} {unidade} {objeto}".lower()
                    
                    # Se tiver QUALQUER indicativo de medicamento/farmácia
                    if any(termo in texto_analise for termo in termos_medicamentos):
                        cnpj_fornecedor = item.get('niFornecedor')
                        
                        registro = {
                            "Órgão Comprador": orgao,
                            "Unidade / Hospital": unidade,
                            "CNPJ Órgão": item.get('orgaoEntidade', {}).get('cnpj'),
                            "UF Órgão": item.get('unidadeOrgao', {}).get('ufSigla'),
                            "Objeto / Descrição da Compra": objeto,
                            "Valor Total Estimado (R$)": item.get('valorTotalEstimado'),
                            "Modalidade": item.get('modalidadeNome'),
                            "Data Publicação": item.get('dataPublicacaoPncp'),
                            "Razão Social Fornecedor": item.get('nomeRazaoSocialFornecedor', 'Não informado'),
                            "CNPJ Fornecedor": cnpj_fornecedor
                        }
                        resultados.append(registro)
                        
            except Exception as e:
                print(f"   [!] Erro na requisição ao PNCP: {e}")
                break
                
    print(f"\n[✓] Total de contratações/fornecedores de medicamentos encontrados: {len(resultados)}")
    return resultados


def enriquecer_e_gerar_excel(lista_contratacoes, arquivo_saida="fornecedores_inca_saude.xlsx"):
    if not lista_contratacoes:
        print("\n[-] Nenhum registro localizado.")
        return
        
    df = pd.DataFrame(lista_contratacoes)
    
    # Remove compras repetidas da mesma empresa para a mesma descrição
    df = df.drop_duplicates(subset=["CNPJ Fornecedor", "Objeto / Descrição da Compra"], keep="first")
    
    cnpjs_unicos = [c for c in df['CNPJ Fornecedor'].dropna().unique() if len(re.sub(r'\D', '', str(c))) == 14]
    
    print(f"\n[+] Buscando contatos de e-mail e telefone de {len(cnpjs_unicos)} distribuidores e laboratórios...")
    
    mapa_contatos = {}
    for idx, cnpj in enumerate(cnpjs_unicos, start=1):
        print(f"    ({idx}/{len(cnpjs_unicos)}) Enriquecendo CNPJ: {cnpj}")
        mapa_contatos[cnpj] = consultar_dados_cnpj(cnpj)
        time.sleep(0.2)
        
    df["E-mail Comercial"] = df["CNPJ Fornecedor"].map(lambda c: mapa_contatos.get(c, {}).get("E-mail Comercial", ""))
    df["Telefone Comercial"] = df["CNPJ Fornecedor"].map(lambda c: mapa_contatos.get(c, {}).get("Telefone Comercial", ""))
    df["UF Fornecedor"] = df["CNPJ Fornecedor"].map(lambda c: mapa_contatos.get(c, {}).get("UF Fornecedor", ""))
    df["Município Fornecedor"] = df["CNPJ Fornecedor"].map(lambda c: mapa_contatos.get(c, {}).get("Município Fornecedor", ""))
    df["Situação Cadastral"] = df["CNPJ Fornecedor"].map(lambda c: mapa_contatos.get(c, {}).get("Situação Cadastral", ""))
    
    colunas_finais = [
        "Razão Social Fornecedor",
        "CNPJ Fornecedor",
        "E-mail Comercial",
        "Telefone Comercial",
        "UF Fornecedor",
        "Município Fornecedor",
        "Situação Cadastral",
        "Órgão Comprador",
        "Unidade / Hospital",
        "UF Órgão",
        "Objeto / Descrição da Compra",
        "Valor Total Estimado (R$)",
        "Modalidade",
        "Data Publicação"
    ]
    
    df = df.reindex(columns=colunas_finais)
    df.to_excel(arquivo_saida, index=False)
    print(f"\n[✓] Sucesso! Base nacional de medicamentos e fornecedores salva em: '{arquivo_saida}'")


if __name__ == "__main__":
    # Varredura nos dados de publicação recentes
    DATA_INICIAL = "20250101"
    DATA_FINAL = "20261231"
    
    contratacoes = buscar_todos_medicamentos_brasil(
        data_inicio=DATA_INICIAL,
        data_fim=DATA_FINAL,
        paginas_max=20
    )
    
    enriquecer_e_gerar_excel(contratacoes, arquivo_saida="fornecedores_inca_saude.xlsx")
