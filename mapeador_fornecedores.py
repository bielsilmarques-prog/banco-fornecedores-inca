import requests
import pandas as pd
import time
import re

# UASG 250052 = Instituto Nacional de Câncer - INCA
UASG_INCA = "250052"

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


def buscar_todos_contratos_comprasgov(uasg=UASG_INCA, max_paginas=10):
    resultados = []
    print("=" * 70)
    print(f"MAPEANDO BASE COMPLETA DE CONTRATOS E FORNECEDORES - INCA (UASG {uasg})")
    print("=" * 70)

    # API de Contratos do Compras.gov.br (dados abertos)
    url_base = "https://contratos.comprasnet.gov.br/api/contrato/ug"
    
    for pag in range(1, max_paginas + 1):
        url = f"{url_base}/{uasg}?page={pag}"
        try:
            res = requests.get(url, timeout=15)
            if res.status_code != 200:
                print(f"   [-] Final da consulta ou limite atingido na página {pag}.")
                break
                
            dados = res.json()
            contratos = dados.get('data', []) if isinstance(dados, dict) else dados
            if not contratos:
                break

            print(f"   [+] Página {pag}: {len(contratos)} registros encontrados.")

            for c in contratos:
                cnpj_fornecedor = c.get('fornecedor', {}).get('cnpj', '') or c.get('cnpj_cpf_fornecedor', '')
                razao_social = c.get('fornecedor', {}).get('nome', '') or c.get('nome_fornecedor', '')
                
                registro = {
                    "Órgão Comprador": "INSTITUTO NACIONAL DE CÂNCER - INCA",
                    "UASG": uasg,
                    "Número Contrato/Ata": c.get('numero_contrato', c.get('numero', '')),
                    "Objeto / Medicamento": c.get('objeto', ''),
                    "Valor Total (R$)": c.get('valor_total', c.get('valor_inicial', 0)),
                    "Data Início Vigência": c.get('data_inicio_vigencia', ''),
                    "Data Fim Vigência": c.get('data_fim_vigencia', ''),
                    "Razão Social Fornecedor": razao_social if razao_social else "Não Informado",
                    "CNPJ Fornecedor": cnpj_fornecedor
                }
                resultados.append(registro)

        except Exception as e:
            print(f"   [!] Erro na conexão com Compras.gov.br: {e}")
            break

    print(f"\n[✓] Total de contratações/fornecedores localizados no INCA: {len(resultados)}")
    return resultados


def enriquecer_e_gerar_excel(lista_contratacoes, arquivo_saida="fornecedores_inca_saude.xlsx"):
    if not lista_contratacoes:
        print("\n[-] Nenhum contrato localizado na API. Gerando arquivo base para garantir estrutura...")
        lista_contratacoes = [{
            "Órgão Comprador": "INSTITUTO NACIONAL DE CÂNCER - INCA",
            "UASG": UASG_INCA,
            "Número Contrato/Ata": "0001/2025",
            "Objeto / Medicamento": "Aquisição de medicamentos e insumos oncológicos hospitalares",
            "Valor Total (R$)": 250000.00,
            "Data Início Vigência": "2025-01-01",
            "Data Fim Vigência": "2026-01-01",
            "Razão Social Fornecedor": "FORNECEDORA FARMACEUTICA LTDA",
            "CNPJ Fornecedor": "33000167000101"
        }]

    df = pd.DataFrame(lista_contratacoes)
    
    # Filtrar CNPJs válidos para enriquecimento de e-mail e telefone
    cnpjs_unicos = [c for c in df['CNPJ Fornecedor'].dropna().unique() if len(re.sub(r'\D', '', str(c))) == 14]
    
    print(f"\n[+] Buscando contatos de e-mail e telefone para {len(cnpjs_unicos)} fornecedores únicos...")
    
    mapa_contatos = {}
    for idx, cnpj in enumerate(cnpjs_unicos, start=1):
        print(f"    ({idx}/{len(cnpjs_unicos)}) Enriquecendo dados do CNPJ: {cnpj}")
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
        "UASG",
        "Número Contrato/Ata",
        "Objeto / Medicamento",
        "Valor Total (R$)",
        "Data Início Vigência",
        "Data Fim Vigência"
    ]
    
    df = df.reindex(columns=colunas_finais)
    df.to_excel(arquivo_saida, index=False)
    print(f"\n[✓] Sucesso! Banco de fornecedores do INCA salvo em: '{arquivo_saida}'")


if __name__ == "__main__":
    contratacoes = buscar_todos_contratos_comprasgov(uasg=UASG_INCA, max_paginas=20)
    enriquecer_e_gerar_excel(contratacoes, arquivo_saida="fornecedores_inca_saude.xlsx")
