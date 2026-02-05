import pandas as pd
import unicodedata
import os
import re
from google import genai
from time import sleep

# =================================================================
# CONFIGURAÇÕES
# =================================================================
ARQUIVO_PRODUTOS = "Produtos_Sem_NCM4.xlsx" 
ARQUIVO_NCM_REF = "Tabela_NCM_Vigente.xlsx"
ARQUIVO_SAIDA = "Planilha_NCM_Final4.xlsx" 
ARQUIVO_PENDENCIAS = "Produtos_Sem_NCM5.xlsx"

API_KEY = os.environ.get("API_KEY")
MODEL_NAME = "gemini-2.0-flash"

client = genai.Client(api_key=API_KEY)

def normalizar(texto):
    if pd.isna(texto): return ""
    nfkd = unicodedata.normalize("NFD", str(texto))
    texto_limpo = "".join([c for c in nfkd if not unicodedata.category(c) == "Mn"])
    return " ".join(texto_limpo.upper().split())

def limpar_ncm(ncm):
    if pd.isna(ncm): return ""
    return re.sub(r"\D", "", str(ncm))

def buscar_ncm_gemini(produto, contexto_tabela):
    prompt = f"""
Você é um auditor fiscal. Classifique o PRODUTO com o NCM de 8 dígitos correto.
PRODUTO: {produto}
EXEMPLOS DA TABELA OFICIAL:
{contexto_tabela.to_string(index=False)}
Responda APENAS o número do NCM (8 dígitos). Se não tiver certeza, responda '0'.
"""
    try:
        response = client.models.generate_content(model=MODEL_NAME, contents=prompt)
        res = re.sub(r"\D", "", response.text.strip())
        return res if len(res) >= 6 else None
    except: return None

# --- 1. CARGA ---
print("📂 Carregando Tabela Vigente...")
df_ref = pd.read_excel(ARQUIVO_NCM_REF, dtype=str)
df_ref.columns = [str(c).strip().upper() for c in df_ref.columns]

col_cod_ref = next((c for c in df_ref.columns if "COD" in c or "NCM" in c), None)
col_desc_ref = next((c for c in df_ref.columns if "DESCR" in c or "NOME" in c), None)

print("🧠 Criando base de conhecimento fiscal...")
pares_validos = set()
ncms_existentes = set()
mapa_oficial = {}

for _, row in df_ref.iterrows():
    cod = limpar_ncm(row[col_cod_ref])
    desc = normalizar(row[col_desc_ref])
    if cod and desc:
        pares_validos.add((desc, cod))
        ncms_existentes.add(cod)
        mapa_oficial[desc] = cod

# --- 2. AUDITORIA ---
print(f"📂 Lendo seus produtos: {ARQUIVO_PRODUTOS}")
df_prod = pd.read_excel(ARQUIVO_PRODUTOS, dtype=str)
df_prod.columns = [str(c).strip().upper() for c in df_prod.columns]

col_nome_prod = next((c for c in df_prod.columns if "NOME" in c or "PRODUTO" in c or "DESCR" in c), None)
if "NCM" not in df_prod.columns: df_prod["NCM"] = ""

pendentes = []
print(f"🔍 Auditando {len(df_prod)} itens...")



for i, row in df_prod.iterrows():
    nome_original = row[col_nome_prod]
    nome_f = normalizar(nome_original)
    ncm_at = limpar_ncm(row.get("NCM", ""))
    ncm_final = ""

    # Validação
    if ncm_at and (nome_f, ncm_at) in pares_validos:
        ncm_final = ncm_at
    else:
        if nome_f in mapa_oficial:
            ncm_final = mapa_oficial[nome_f]
        else:
            pals = nome_f.split()[:2]
            subset = df_ref[df_ref[col_desc_ref].str.contains('|'.join(pals), na=False, case=False)].head(10)
            ncm_ia = buscar_ncm_gemini(nome_original, subset)
            
            if ncm_ia in ncms_existentes:
                ncm_final = ncm_ia
                print(f"✅ IA encontrou: {nome_original} -> {ncm_ia}")
                sleep(1.2)
            else:
                # MENSAGEM DE ALERTA PARA NÃO ENCONTRADO
                print(f"❌ NÃO ENCONTRADO: {nome_original}")

    df_prod.at[i, "NCM"] = ncm_final
    if not ncm_final:
        pendentes.append(row)

# --- 3. SALVAMENTO ---
df_prod.to_excel(ARQUIVO_SAIDA, index=False)
if pendentes:
    pd.DataFrame(pendentes).to_excel(ARQUIVO_PENDENCIAS, index=False)
    print(f"\n⚠️ Planilha de pendências gerada com {len(pendentes)} itens: {ARQUIVO_PENDENCIAS}")

print(f"\n✅ PROCESSO CONCLUÍDO!")
print(f"💾 Arquivo final: {ARQUIVO_SAIDA}")