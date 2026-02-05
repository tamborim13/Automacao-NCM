import pandas as pd
import unicodedata
import os
import re

# =========================
# CONFIGURAÇÕES
# =========================
ARQUIVO_PRODUTOS = "Produtos_Sem_NCM.xlsx"
ARQUIVO_NCM_VIGENTE = "Tabela_NCM_Vigente.xlsx"
OUTPUT_FILE = "Planilha_NCM_Final.xlsx"

def limpar_texto(texto):
    """Padroniza o texto: remove acentos, espaços extras e deixa em maiúsculo"""
    if pd.isna(texto):
        return ""
    nfkd_form = unicodedata.normalize('NFKD', str(texto))
    texto_limpo = "".join([c for c in nfkd_form if not unicodedata.category(c) == 'Mn'])
    return " ".join(texto_limpo.upper().split())

def limpar_ncm(ncm):
    """Remove qualquer caractere que não seja número"""
    if pd.isna(ncm):
        return ""
    return re.sub(r"\D", "", str(ncm))

# =========================
# CARGA E MAPEAMENTO
# =========================

print("📂 Carregando Tabela Vigente e criando base de validação...")
df_vigente = pd.read_excel(ARQUIVO_NCM_VIGENTE, dtype=str)
df_vigente.columns = [limpar_texto(c) for c in df_vigente.columns]

# Identificar colunas na Tabela Vigente
col_descr_ref = next((c for c in df_vigente.columns if "DESCR" in c), None)
col_ncm_ref = next((c for c in df_vigente.columns if "COD" in c or "NCM" in c), None)

if not col_descr_ref or not col_ncm_ref:
    raise ValueError("❌ Colunas de Código ou Descrição não encontradas na Tabela Vigente.")

# Criar set de NCMs válidos e mapa de Descrição -> NCM
ncm_validos = set()
mapa_descricao_ncm = {}

for _, row in df_vigente.iterrows():
    ncm_oficial = limpar_ncm(row[col_ncm_ref])
    desc_oficial = limpar_texto(row[col_descr_ref])
    
    if ncm_oficial:
        ncm_validos.add(ncm_oficial)
    if desc_oficial and ncm_oficial:
        mapa_descricao_ncm[desc_oficial] = ncm_oficial

print(f"✅ Base oficial: {len(ncm_validos)} NCMs e {len(mapa_descricao_ncm)} descrições mapeadas.")

# =========================
# PROCESSAMENTO DOS PRODUTOS
# =========================

print(f"📂 Carregando {ARQUIVO_PRODUTOS}...")
df_prod = pd.read_excel(ARQUIVO_PRODUTOS, dtype=str)
df_prod.columns = [limpar_texto(c) for c in df_prod.columns]

# Identificar coluna de nome nos produtos
col_nome_prod = next((c for c in df_prod.columns if "NOME" in c or "DESCR" in c), None)

print("🔍 Iniciando validação e batimento exato...")

def processar_linha(linha):
    nome_original = linha[col_nome_prod]
    nome_limpo = limpar_texto(nome_original)
    ncm_atual = limpar_ncm(linha.get("NCM", ""))

    # 1. TESTE DE VALIDADE: Se já tem NCM, ele existe na tabela oficial?
    if ncm_atual and ncm_atual in ncm_validos:
        return ncm_atual # NCM é válido, mantém.
    
    # 2. TESTE DE DESCRIÇÃO: Se NCM era inválido ou estava vazio, a descrição bate 100%?
    if nome_limpo in mapa_descricao_ncm:
        return mapa_descricao_ncm[nome_limpo] # Achou por descrição exata.
    
    # 3. FALHA: Se não passou em nenhum, deixa em branco.
    return ""

# Aplicar lógica
df_prod['NCM'] = df_prod.apply(processar_linha, axis=1)

# =========================
# FINALIZAÇÃO
# =========================

df_prod.to_excel(OUTPUT_FILE, index=False)

total = len(df_prod)
preenchidos = df_prod[df_prod['NCM'] != ""].shape[0]

print(f"\n✅ PROCESSO CONCLUÍDO!")
print(f"📊 Total de produtos: {total}")
print(f"✅ NCMs Validados/Encontrados: {preenchidos}")
print(f"⚠️ NCMs em branco (Inválidos ou Sem Batimento): {total - preenchidos}")
print(f"💾 Arquivo salvo em: {OUTPUT_FILE}")