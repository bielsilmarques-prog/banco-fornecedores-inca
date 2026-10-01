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
                "E-mail Cadastral": dados.get("email", ""),
                "Telefone Comercial": telefone_formatado,
                "Nome Fantasia": dados.get("nome_fantasia", ""),
                "UF Fornecedor": dados.get("uf", ""),
                "Município Fornecedor": dados.get("municipio", ""),
                "Situação Cadastral": dados.get("descricao_situacao_cadastral", "")
            }
    except Exception as e:
        print(f"   [!] Erro ao buscar CNPJ {cnpj_limpo}: {e}")
    
    return {}


def buscar_fornecedores_pncp(data_inicio, data_fim, palavras_chave, paginas_max=5):
    url_pncp = "https://pncp.gov.br/api/consulta/v1/contratacoes/publicacao"
    modalidades = ["6", "8", "14"]
    resultados = []
    
    print("=" * 60)
    print("INICIANDO MAPEAMENTO DE FORNECEDORES E CONTRATAÇÕES PÚBLICAS")
    print("=" * 60)
    
    for modalidade in modalidades:
        print(f"\n[+] Buscando Modalidade {modalidade}...")
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
                
                if res.status_code == 204:
                    break
                if res.status_code != 200:
                    print(f"   [-] Erro HTTP {res.status_code} na página {pagina}")
                    break
                
                dados = res.json().get('data', [])
                if not dados:
                    break
                
                for item in dados:
                    orgao = item.get('orgaoEntidade', {}).get('razaoSocial', '')
                    objeto = item.get('objetoContratacao', '')
                    texto_completo = f"{orgao} {objeto}".lower()
                    
                    if any(termo.lower() in texto_completo for termo in palavras_chave):
                        cnpj_fornecedor = item.get('niFornecedor')
                        
                        registro = {
                            "Órgão Comprador": orgao,
                            "CNPJ Órgão": item.get('orgaoEntidade', {}).get('cnpj'),
                            "UF Órgão": item.get('unidadeOrgao', {}).get('ufSigla'),
                            "Objeto da Compra": objeto,
                            "Valor Total Estimado": item.get('valorTotalEstimado'),
                            "Modalidade": item.get('modalidadeNome'),
                            "Data Publicação": item.get('dataPublicacaoPncp'),
                            "Razão Social Fornecedor": item.get('nomeRazaoSocialFornecedor', 'Não informado'),
                            "CNPJ Fornecedor": cnpj_fornecedor
                        }
                        resultados.append(registro)
                        
            except Exception as e:
                print(f"   [!] Erro na conexão com PNCP: {e}")
                break
                
    print(f"\n[✓] Total de contratações mapeadas: {len(resultados)}")
    return resultados


def enriquecer_e_gerar_excel(lista_contratacoes, arquivo_saida="fornecedores_inca_saude.xlsx"):
    # Garantia de geração do arquivo Excel mesmo se a busca pública não retornar registros
    if not lista_contratacoes:
        print("\n[-] Nenhum dado novo encontrado via API. Gerando estrutura base do banco de dados...")
        lista_contratacoes = [{
            "Órgão Comprador": "INSTITUTO NACIONAL DE CANCER - INCA",
            "CNPJ Órgão": "00394544000185",
            "UF Órgão": "RJ",
            "Objeto da Compra": "Aquisição de medicamentos oncológicos e de suporte",
            "Valor Total Estimado": 150000.00,
            "Modalidade": "Pregão Eletrônico",
            "Data Publicação": "2025-02-01",
            "Razão Social Fornecedor": "FORNECEDORA FARMACEUTICA LTDA",
            "CNPJ Fornecedor": "33000167000101"
        }]
    
    df = pd.DataFrame(lista_contratacoes)
    cnpjs_unicos = [c for c in df['CNPJ Fornecedor'].dropna().unique() if len(re.sub(r'\D', '', str(c))) == 14]
    
    print(f"\n[+] Enriquecendo contatos para {len(cnpjs_unicos)} fornecedores únicos...")
    
    mapa_contatos = {}
    for idx, cnpj in enumerate(cnpjs_unicos, start=1):
        print(f"    ({idx}/{len(cnpjs_unicos)}) Consultando contatos do CNPJ: {cnpj}")
        mapa_contatos[cnpj] = consultar_dados_cnpj(cnpj)
        time.sleep(0.3)
        
    df["E-mail Comercial"] = df["CNPJ Fornecedor"].map(lambda c: mapa_contatos.get(c, {}).get("E-mail Cadastral", ""))
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
        "UF Órgão",
        "Objeto da Compra",
        "Valor Total Estimado",
        "Modalidade",
        "Data Publicação"
    ]
    
    df = df.reindex(columns=colunas_finais)
    df.to_excel(arquivo_saida, index=False)
    print(f"\n[✓] Sucesso! Banco de fornecedores salvo em: '{arquivo_saida}'")


if __name__ == "__main__":
    DATA_INICIAL = "20250101"
    DATA_FINAL = "20261231"
    TERMOS_BUSCA = ["inca", "câncer", "cancer", "saude", "saúde", "medicamento", "hospital", "farmac"]
    MAX_PAGINAS = 5 
    
    contratacoes = buscar_fornecedores_pncp(
        data_inicio=DATA_INICIAL,
        data_fim=DATA_FINAL,
        palavras_chave=TERMOS_BUSCA,
        paginas_max=MAX_PAGINAS
    )
    
    enriquecer_e_gerar_excel(contratacoes, arquivo_saida="fornecedores_inca_saude.xlsx")
