import pandas as pd
import os

# --- CONFIGURAÇÕES ---
ARQUIVO_PRINCIPAL = "Planilha_Final_base.xlsx"  # Sua planilha original
ARQUIVO_PREENCHIDO = "Planilha_NCM_Final.xlsx" # A que você preencheu agora
ARQUIVO_SAIDA = "Planilha_Final_base5.xlsx"

# 1. Carregar as planilhas
# Forçamos o NCM como string para não perder o zero à esquerda (ex: 0101.21.00)
df_principal = pd.read_excel(ARQUIVO_PRINCIPAL, dtype={'NCM': str})
df_preenchido = pd.read_excel(ARQUIVO_PREENCHIDO, dtype={'NCM': str})

# 2. Normalizar os nomes para garantir que o "PRODUTO A" case com "produto a"
df_principal['NOME_NORM'] = df_principal['NOME'].astype(str).str.strip().str.upper()
df_preenchido['NOME_NORM'] = df_preenchido['NOME'].astype(str).str.strip().str.upper()

# 3. Criar um dicionário de consulta (De: Nome -> Para: NCM)
# Removemos duplicatas da planilha preenchida para não dar erro
mapa_ncm = df_preenchido.drop_duplicates('NOME_NORM').set_index('NOME_NORM')['NCM'].to_dict()

# 4. Função de preenchimento inteligente
def atualizar_ncm(linha):
    # Se o NCM já existe na principal, mantém ele
    if pd.notna(linha['NCM']) and str(linha['NCM']).strip() != "":
        return linha['NCM']
    
    # Se estiver vazio, busca no nosso dicionário da planilha preenchida
    return mapa_ncm.get(linha['NOME_NORM'], linha['NCM'])

# 5. Aplicar a atualização
df_principal['NCM'] = df_principal.apply(atualizar_ncm, axis=1)

# 6. Remover coluna auxiliar e salvar
df_principal = df_principal.drop(columns=['NOME_NORM'])
df_principal.to_excel(ARQUIVO_SAIDA, index=False)

print(f"✅ Processo concluído! Os NCMs foram mesclados em: {ARQUIVO_SAIDA}")