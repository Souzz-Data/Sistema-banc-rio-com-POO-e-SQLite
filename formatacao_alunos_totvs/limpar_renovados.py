import pandas as pd
import re
import os
import sys
import unicodedata
import warnings

warnings.filterwarnings('ignore')

# ==========================================
# 1. DEFINIÇÃO DE CAMINHOS FIXOS
# ==========================================
# Uso: python limpar_renovados.py [base_disparo.xlsx] [forms_renovacao.xlsx] [base_totvs.xlsx] [pasta_saida]
arquivo_disparo = r"C:\Users\everton.araujo\Desktop\Disparo_Limpeza\BaseDisparo_Alunos_Ativos.xlsx"
arquivo_forms = r"C:\Users\everton.araujo\Desktop\Disparo_Limpeza\Registro - Renovação (respostas).xlsx"
arquivo_totvs = r"C:\Users\everton.araujo\Desktop\Disparo_Limpeza\Base_Ativos\lista_alunos.xlsx"
caminho_output = r"C:\Users\everton.araujo\Desktop\Disparo_Limpeza"

argumentos = sys.argv[1:]
if len(argumentos) > 0: arquivo_disparo = argumentos[0]
if len(argumentos) > 1: arquivo_forms = argumentos[1]
if len(argumentos) > 2: arquivo_totvs = argumentos[2]
if len(argumentos) > 3: caminho_output = argumentos[3]

if not os.path.exists(caminho_output):
    os.makedirs(caminho_output)

# Filiais que não entram no disparo (coluna Tags)
FILIAIS_EXCLUIDAS = ['VILA MILITAR']

COL_FORMS_NOME = 'NOME DO ALUNO:'
COL_FORMS_RA = 'RA:'

# ==========================================
# 2. FUNÇÕES DE LIMPEZA
# ==========================================
def limpar_nome(nome):
    if pd.isna(nome) or str(nome).strip() == '':
        return ''
    nome_str = str(nome).strip().upper()
    if '@' in nome_str and not nome_str.startswith('@'):
        nome_str = nome_str.split('@')[0].replace('.', ' ').replace('_', ' ')
    nome_str = re.sub(r'^LEAD\s*-\s*', '', nome_str)
    nome_str = re.sub(r'\s*\(\d+\)\s*', '', nome_str)
    nome_str = re.sub(r'[^A-ZÁÀÂÃÉÈÊÍÏÓÔÕÖÚÇÑ\s]', '', nome_str)
    return re.sub(r'\s+', ' ', nome_str).strip()

def nome_comparacao(nome):
    # Sem acento, para casar "Mateus Medina" do Forms com "MATEUS MEDINA" do TOTVS
    nome_str = limpar_nome(nome)
    nome_str = unicodedata.normalize('NFKD', nome_str).encode('ascii', 'ignore').decode()
    return re.sub(r'\s+', ' ', nome_str).strip()

def limpar_ra(ra):
    return re.sub(r'\D', '', str(ra)) if pd.notna(ra) else ''

# ==========================================
# 3. LEITURA DAS BASES
# ==========================================
print("Lendo base de disparo, Forms de renovação e base do TOTVS...")
try:
    df_disparo = pd.read_excel(arquivo_disparo, dtype=str)
    df_forms = pd.read_excel(arquivo_forms, dtype=str)
    df_totvs = pd.read_excel(arquivo_totvs, dtype=str)
except FileNotFoundError as e:
    print(f"❌ Erro ao ler arquivos: {e}")
    exit()

qtd_original = len(df_disparo)
df_disparo['Nome_Comparacao'] = df_disparo['Nome'].apply(nome_comparacao)

# ==========================================
# 4. IDENTIFICA OS RENOVADOS (RA -> NOME NO TOTVS)
# ==========================================
df_totvs['RA'] = df_totvs['NOME'].astype(str).str.extract(r'\((\d+)\)')[0]
df_totvs['Nome_Comparacao'] = df_totvs['NOME'].apply(nome_comparacao)
nome_por_ra = df_totvs.dropna(subset=['RA']).drop_duplicates('RA').set_index('RA')['Nome_Comparacao']

df_forms['RA_Limpo'] = df_forms[COL_FORMS_RA].apply(limpar_ra)
df_forms['Nome_Forms'] = df_forms[COL_FORMS_NOME].apply(nome_comparacao)

# Pelo RA usa o nome do TOTVS; o nome digitado no Forms também conta (RA digitado errado)
df_forms['Nome_TOTVS'] = df_forms['RA_Limpo'].map(nome_por_ra).fillna('')
nomes_renovados = (set(df_forms['Nome_TOTVS']) | set(df_forms['Nome_Forms'])) - {''}

# ==========================================
# 5. CORTES NA BASE DE DISPARO
# ==========================================
renovado = df_disparo['Nome_Comparacao'].isin(nomes_renovados)
df_limpo = df_disparo[~renovado]
qtd_renovados = renovado.sum()

filial_excluida = df_limpo['Tags'].astype(str).str.strip().str.upper().isin(FILIAIS_EXCLUIDAS)
df_limpo = df_limpo[~filial_excluida]
qtd_filiais = filial_excluida.sum()

# Respostas do Forms que não acharam aluno na base de disparo (para conferência)
nomes_no_disparo = set(df_disparo['Nome_Comparacao'])
encontrado = df_forms['Nome_TOTVS'].isin(nomes_no_disparo) | df_forms['Nome_Forms'].isin(nomes_no_disparo)
df_nao_encontrados = df_forms[~encontrado].drop(columns=['RA_Limpo', 'Nome_Forms', 'Nome_TOTVS'])
df_nao_encontrados['Motivo'] = df_forms.loc[~encontrado, 'Nome_TOTVS'].map(
    lambda n: 'Está no TOTVS, mas não estava no disparo (sem celular ou duplicado)' if n else 'RA/Nome não encontrado no TOTVS'
)

# ==========================================
# 6. EXPORTAÇÃO E RELATÓRIO
# ==========================================
colunas_finais = ['Nome', 'Telefone', 'country code', 'Tags', 'E-mail']
df_final = df_limpo[colunas_finais]

nome_base = os.path.splitext(os.path.basename(arquivo_disparo))[0]
arquivo_saida = os.path.join(caminho_output, f"{nome_base}_SemRenovados.xlsx")
df_final.to_excel(arquivo_saida, index=False)

arquivo_conferencia = os.path.join(caminho_output, "Renovados_Nao_Encontrados.xlsx")
df_nao_encontrados.to_excel(arquivo_conferencia, index=False)

print("\n" + "="*60)
print(f"✅ SUCESSO! Renovados e filiais excluídas retirados da base de disparo.")
print(f"📁 Salvo em: {arquivo_saida}")
print(f"📁 Conferência: {arquivo_conferencia}")
print("="*60)

print("\n📊 RELATÓRIO DE EXCLUSÕES (O FUNIL):")
print("-" * 60)
print(f"📉 Total Original: {qtd_original} contatos")
print(f"📝 Respostas no Forms de renovação: {len(df_forms)} ({df_forms['RA_Limpo'].replace('', pd.NA).nunique()} RAs)")
print(f"🔁 Cortados por já terem RENOVADO: -{qtd_renovados}")
print(f"🏢 Cortados por filial excluída ({', '.join(FILIAIS_EXCLUIDAS)}): -{qtd_filiais}")
print(f"🔎 Renovados que não estavam no disparo: {len(df_nao_encontrados)}")
print("-" * 60)
print(f"🚀 TOTAL FINAL PRONTO PARA DISPARO: {len(df_final)} contatos")
print("="*60)

print("\n🏢 RESUMO POR FILIAL (TAGS):")
print("-" * 60)
for filial, quantidade in df_final['Tags'].value_counts().items():
    print(f"🔹 {filial}: {quantidade} contato(s)")
print("-" * 60)
