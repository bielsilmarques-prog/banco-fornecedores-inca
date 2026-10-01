import requests
import pandas as pd
import time
import re

CNPJ_MINISTERIO_SAUDE = "00394544000185"

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


def buscar_todas_contratacoes_inca(data_inicio, data_fim, max_paginas=20):
    resultados = []
    print("=" * 70)
    print("INICIANDO BUSCA COMPLETA DE FORNECEDORES DO INCA / MINISTÉRIO DA SAÚDE")
    print("=" * 70)

    # 1. Busca por órgão direto no PNCP
    url_orgao = f"https://pncp.gov.br/api/pncp/v1/orgaos/{CNPJ_MINISTERIO_SAUDE}/contratacoes/publicacao"
    
    for pagina in range(1, max_paginas + 1):
        params = {
            "dataInicial": data_inicio,
            "dataFinal": data_fim,
            "pagina": pagina,
            "tam_pagina": 50
        }
        try:
            res = requests.get(url_orgao, params=params, timeout=15)
            if res.status_code in [204, 404]:
                break
            if res.status_code != 200:
                print(f"   [-] Status {res.status_code} na página {pagina}")
                break
                
            dados = res.json()
            itens = dados.get('data', []) if isinstance(dados, dict) else dados
            if not itens:
                break

            for item in itens:
                orgao = item.get('orgaoEntidade', {}).get('razaoSocial', '')
                unidade = item.get('unidadeOrgao', {}).get('nomeUnidade', '')
                objeto = item.get('objetoContratacao', '')
                texto = f"{orgao} {unidade} {objeto}".lower()
                
                # Filtra especificamente unidades do INCA ou compras de saúde/medicamentos
                if "inca" in texto or "cancer" in texto or "câncer" in texto or "medicament" in texto or "farmac" in texto:
                    cnpj_fornecedor = item.get('niFornecedor')
                    registro = {
                        "Órgão Comprador": orgao,
                        "Unidade Compradora": unidade,
                        "CNPJ Órgão": CNPJ_MINISTERIO_SAUDE,
                        "UF Órgão": item.get('unidadeOrgao', {}).get('ufSigla', 'RJ'),
                        "Objeto da Compra": objeto,
                        "Valor Total Estimado": item.get('valorTotalEstimado'),
                        "Modalidade": item.get('modalidadeNome'),
                        "Data Publicação": item.get('dataPublicacaoPncp'),
                        "Razão Social Fornecedor": item.get('nomeRazaoSocialFornecedor', 'Não informado'),
                        "CNPJ Fornecedor": cnpj_fornecedor
                    }
                    resultados.append(registro)

        except Exception as e:
            print(f"   [!] Erro na requisição: {e}")
            break

    # 2. Busca genérica de apoio no endpoint público de contratações se a busca por órgão for restrita
    if len(resultados) < 5:
        print("\n[+] Ampliando busca por palavras-chave globais no PNCP...")
        url_busca = "https://pncp.gov.br/api/consulta/v1/contratacoes/publicacao"
        modalidades = ["6", "8", "14"]
        
        for mod in modalidades:
            for pag in range(1, 10):
                params_g = {
                    "dataInicial": data_inicio,
                    "dataFinal": data_fim,
                    "codigoModalidadeContratacao": mod,
                    "pagina": pag,
                    "tamanhoPagina": 50
                }
                try:
                    res = requests.get(url_busca, params=params_g, timeout=15)
                    if res.status_code != 200 or not res.json().get('data'):
                        break
                    
                    for item in res.json().get('data', []):
                        orgao = item.get('orgaoEntidade', {}).get('razaoSocial', '')
                        objeto = item.get('objetoContratacao', '')
                        if "inca" in f"{orgao} {objeto}".lower() or "cancer" in f"{orgao} {objeto}".lower():
                            resultados.append({
                                "Órgão Comprador": orgao,
                                "Unidade Compradora": item.get('unidadeOrgao', {}).get('nomeUnidade', ''),
                                "CNPJ Órgão": item.get('orgaoEntidade', {}).get('cnpj'),
                                "UF Órgão": item.get('unidadeOrgao', {}).get('ufSigla'),
                                "Objeto da Compra": objeto,
                                "Valor Total Estimado": item.get('valorTotalEstimado'),
                                "Modalidade": item.get('modalidadeNome'),
                                "Data Publicação": item.get('dataPublicacaoPncp'),
                                "Razão Social Fornecedor": item.get('nomeRazaoSocialFornecedor', 'Não informado'),
                                "CNPJ Fornecedor": item.get('niFornecedor')
                            })
                except Exception:
                    break

    print(f"\n[✓] Total de contratações mapeadas: {len(resultados)}")
    return resultados


def enriquecer_e_gerar_excel(lista_contratacoes, arquivo_saida="fornecedores_inca_saude.xlsx"):
    if not lista_contratacoes:
        print("\n[-] Nenhum dado encontrado no período especificado.")
        return
        
    df = pd.DataFrame(lista_contratacoes)
    
    # Remove duplicados de contratações idênticas
    df = df.drop_duplicates(subset=["CNPJ Fornecedor", "Objeto da Compra"], keep="first")
    
    cnpjs_unicos = [c for c in df['CNPJ Fornecedor'].dropna().unique() if len(re.sub(r'\D', '', str(c))) == 14]
    
    print(f"\n[+] Enriquecendo dados cadastrais de {len(cnpjs_unicos)} fornecedores únicos...")
    
    mapa_contatos = {}
    for idx, cnpj in enumerate(cnpjs_unicos, start=1):
        print(f"    ({idx}/{len(cnpjs_unicos)}) Consultando CNPJ: {cnpj}")
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
        "Unidade Compradora",
        "Objeto da Compra",
        "Valor Total Estimado",
        "Modalidade",
        "Data Publicação"
    ]
    
    df = df.reindex(columns=colunas_finais)
    df.to_excel(arquivo_saida, index=False)
    print(f"\n[✓] Sucesso! Planilha gerada com {len(df)} fornecedores/compras em: '{arquivo_saida}'")


if __name__ == "__main__":
    # Ampliado o período de consulta para pegar o histórico acumulado recente
    DATA_INICIAL = "20240101"
    DATA_FINAL = "20261231"
    
    contratacoes = buscar_todas_contratacoes_inca(
        data_inicio=DATA_INICIAL,
        data_fim=DATA_FINAL,
        max_paginas=50
    )
    
    enriquecer_e_gerar_excel(contratacoes, arquivo_saida="fornecedores_inca_saude.xlsx")
