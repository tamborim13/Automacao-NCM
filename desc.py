import pandas as pd
import unicodedata
import os
import re

# =========================================================
# CONFIGURAÇÕES
# =========================================================

ARQUIVO_PRODUTOS = "PlanilhaNCM.xlsx"
ARQUIVO_NCM_RECEITA = "Tabela_NCM_Vigente.xlsx"

OUTPUT_FILE = "Planilha_Final.xlsx"
OUTPUT_NAO_ENCONTRADOS = "NCMs_Sem_Descricao.xlsx"

SALVAR_A_CADA = 100

DTYPE_PRODUTOS = {'NCM': str}

# =========================================================
# FUNÇÕES AUXILIARES
# =========================================================

def normalizar(texto):
    if pd.isna(texto):
        return ""
    texto = str(texto).strip()
    return ''.join(
        c for c in unicodedata.normalize('NFD', texto)
        if unicodedata.category(c) != 'Mn'
    ).upper()


def limpar_ncm(valor):
    """Remove tudo que não é número e garante 8 dígitos"""
    if pd.isna(valor):
        return None
    return re.sub(r'\D', '', str(valor)).zfill(8)


def salvar(df):
    df.to_excel(OUTPUT_FILE, index=False)
    print(f"💾 Progresso salvo em {OUTPUT_FILE}")


# =========================================================
# BUSCA DE DESCRIÇÃO COM FALLBACK (8 → 6 → 4 → 2)
# =========================================================

def buscar_descricao_completa(ncm_raw, df_ncm, col_descr):
    ncm = limpar_ncm(ncm_raw)
    if not ncm:
        return None, None

    # 1️⃣ Exato (8 dígitos)
    mask = df_ncm['NCM_BUSCA'] == ncm
    if mask.any():
        return df_ncm.loc[mask, col_descr].iloc[0], "8"

    # 2️⃣ Prefixo 6
    mask = df_ncm['NCM_BUSCA'].str.startswith(ncm[:6])
    if mask.any():
        return df_ncm.loc[mask, col_descr].iloc[0], "6"

    # 3️⃣ Prefixo 4
    mask = df_ncm['NCM_BUSCA'].str.startswith(ncm[:4])
    if mask.any():
        return df_ncm.loc[mask, col_descr].iloc[0], "4"

    # 4️⃣ Prefixo 2 (Capítulo)
    mask = df_ncm['NCM_BUSCA'].str.startswith(ncm[:2])
    if mask.any():
        return df_ncm.loc[mask, col_descr].iloc[0], "2"

    return None, None


# =========================================================
# LEITURA DAS PLANILHAS
# =========================================================

print("📂 Lendo arquivos...")

df_prod = pd.read_excel(
    ARQUIVO_PRODUTOS,
    engine="openpyxl",
    dtype=DTYPE_PRODUTOS
)

df_ncm = pd.read_excel(
    ARQUIVO_NCM_RECEITA,
    engine="openpyxl"
)

# Normaliza colunas
df_prod.columns = [normalizar(c) for c in df_prod.columns]
df_ncm.columns = [normalizar(c) for c in df_ncm.columns]

# Identifica colunas principais da tabela NCM
col_descr = next(c for c in df_ncm.columns if "DESCRICAO" in c)
col_codigo = next(c for c in df_ncm.columns if "CODIGO" in c)

# Cria coluna de busca limpa na tabela NCM
df_ncm['NCM_BUSCA'] = (
    df_ncm[col_codigo]
    .astype(str)
    .str.replace(r'[^\d]', '', regex=True)
    .str.zfill(8)
)

# Garante coluna DESCRICAO no produto
if "DESCRICAO" not in df_prod.columns:
    df_prod["DESCRICAO"] = pd.NA

# =========================================================
# PROCESSAMENTO
# =========================================================

nao_encontrados = []
contador = 0

total = len(df_prod)
print(f"🔍 Processando {total} produtos...\n")

for i, row in df_prod.iterrows():

    if not row.get("NCM") or not str(row["NCM"]).strip():
        continue

    if pd.notna(row["DESCRICAO"]) and str(row["DESCRICAO"]).strip():
        continue

    descricao, nivel = buscar_descricao_completa(
        row["NCM"],
        df_ncm,
        col_descr
    )

    if descricao:
        df_prod.at[i, "DESCRICAO"] = descricao
        print(f"✔ NCM {row['NCM']} → nível {nivel}")
    else:
        nao_encontrados.append(row)
        print(f"✖ NCM {row['NCM']} não encontrado")

    contador += 1
    if contador % SALVAR_A_CADA == 0:
        salvar(df_prod)

# =========================================================
# FINALIZAÇÃO
# =========================================================

salvar(df_prod)

if nao_encontrados:
    df_fail = pd.DataFrame(nao_encontrados)
    df_fail.to_excel(OUTPUT_NAO_ENCONTRADOS, index=False)
    print(f"\n⚠️ {len(df_fail)} NCMs não encontrados salvos em {OUTPUT_NAO_ENCONTRADOS}")

print("\n🚀 PROCESSO CONCLUÍDO COM SUCESSO!")
