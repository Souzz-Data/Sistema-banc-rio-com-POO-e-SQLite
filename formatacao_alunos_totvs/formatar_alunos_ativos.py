import pandas as pd
import numpy as np
import re
import os
import sys
import warnings

warnings.filterwarnings('ignore')

# ==========================================
# 1. DEFINIÇÃO DE CAMINHOS FIXOS
# ==========================================
# Uso: python formatar_alunos_ativos.py [planilha_entrada.xlsx] [pasta_saida]
arquivo_ativos = r"C:\Users\everton.araujo\Desktop\Disparo_Limpeza\Base_Ativos\lista_alunos.xlsx"
caminho_output = r"C:\Users\everton.araujo\Desktop\Disparo_Limpeza"

if len(sys.argv) > 1: arquivo_ativos = sys.argv[1]
if len(sys.argv) > 2: caminho_output = sys.argv[2]

if not os.path.exists(caminho_output):
    os.makedirs(caminho_output)

ABAS_ENTRADA = ['REGULARES', 'MC e TEENS 4']

# Ordem de prioridade: primeiro celular válido encontrado é o usado (fixo só se não houver celular)
COLUNAS_TELEFONE = [
    'TELEFONE 1 ALUNO',
    'TELEFONE 2 ALUNO',
    'TELEFONE 1 RESP. FIN.',
    'TELEFONE 2 RESP. FIN.',
    'TELEFONE COMERCIAL RESP. FIN.',
    'TELEFONE 1 RESP. ACAD.',
    'TELEFONE 2 RESP. ACAD.',
    'TELEFONE 3 RESP. ACAD.',
]
COLUNAS_EMAIL = ['EMAIL', 'EMAIL RESP FIN', 'EMAIL RESP ACA']

# ==========================================
# 2. FUNÇÕES DE LIMPEZA (TELEFONE, NOME E E-MAIL)
# ==========================================
def formatar_telefone(telefone):
    if pd.isna(telefone) or str(telefone).strip() in ['', '(Vazias)', 'nan']:
        return ''
    telefone_str = str(telefone).strip()
    if telefone_str.startswith('+') and not telefone_str.startswith('+55'):
        return ''
    numeros = re.sub(r'\D', '', telefone_str)
    if not numeros: return ''
    if numeros.startswith('0'): numeros = numeros[1:]
    if numeros.startswith('55') and len(numeros) in [12, 13]: pass
    elif len(numeros) in [10, 11]: numeros = '55' + numeros
    elif len(numeros) in [8, 9]: numeros = '5521' + numeros
    else: return ''
    return numeros

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

def is_nome_valido(nome):
    if not nome: return False
    nome_str = str(nome).lower()
    if '{' in nome_str or '}' in nome_str or 'http' in nome_str or 'type:' in nome_str: return False
    if nome_str.startswith('@') or nome_str.startswith('_'): return False
    if re.search(r'\d{4,}', nome_str) or re.search(r'^\(?\d{2}\)?\s*\d+', nome_str): return False
    return True

def limpar_email(email):
    if pd.isna(email):
        return np.nan
    # Pega o pedaço que contém '@' (ex.: "p fulano@gmail.com" -> "fulano@gmail.com")
    for parte in str(email).strip().split():
        if re.fullmatch(r'[^@\s]+@[^@\s]+\.[^@\s]+', parte):
            return parte.lower()
    return np.nan

def is_celular(numero):
    # 55 + DDD + 9 dígitos começando com 9
    return len(numero) == 13 and numero[4] == '9'

def melhor_telefone(linha):
    telefones = [formatar_telefone(linha[col]) for col in COLUNAS_TELEFONE if col in linha.index]
    telefones = [t for t in telefones if t != '']
    if not telefones: return ''
    # Prioriza celular; se não houver, usa o primeiro fixo válido
    celulares = [t for t in telefones if is_celular(t)]
    return celulares[0] if celulares else telefones[0]

def primeiro_valido(linha, colunas, funcao, vazio):
    for col in colunas:
        if col in linha.index:
            valor = funcao(linha[col])
            if not (valor == vazio or (isinstance(valor, float) and pd.isna(valor))):
                return valor
    return vazio

# ==========================================
# 3. LEITURA DA BASE DE ATIVOS (TOTVS)
# ==========================================
print("Lendo Base de Alunos Ativos (TOTVS)...")
try:
    df_ativos = pd.concat(
        [pd.read_excel(arquivo_ativos, sheet_name=aba) for aba in ABAS_ENTRADA],
        ignore_index=True
    )
except (FileNotFoundError, ValueError) as e:
    print(f"❌ Erro ao ler a base de ativos: {e}")
    exit()

qtd_original = len(df_ativos)
qtd_atual = qtd_original

metricas = {
    "Nomes_Invalidos": 0,
    "Sem_Telefone": 0,
    "Duplicados": 0
}

# ==========================================
# 4. LIMPEZA DE FORMATOS
# ==========================================
df_ativos['RA'] = df_ativos['NOME'].astype(str).str.extract(r'\((\d+)\)')[0]
df_ativos['Nome_Limpo'] = df_ativos['NOME'].apply(limpar_nome)
df_ativos = df_ativos[df_ativos['Nome_Limpo'].apply(is_nome_valido)]
df_ativos = df_ativos[df_ativos['Nome_Limpo'] != '']

metricas["Nomes_Invalidos"] = qtd_atual - len(df_ativos)
qtd_atual = len(df_ativos)

df_ativos['Telefone_Limpo'] = df_ativos.apply(melhor_telefone, axis=1)
df_ativos['Email_Limpo'] = df_ativos.apply(lambda l: primeiro_valido(l, COLUNAS_EMAIL, limpar_email, np.nan), axis=1)

df_ativos = df_ativos[df_ativos['Telefone_Limpo'] != '']
metricas["Sem_Telefone"] = qtd_atual - len(df_ativos)
qtd_atual = len(df_ativos)

# ==========================================
# 5. APLICAÇÃO DO MOLDE FILIAIS
# ==========================================
df_molde = pd.DataFrame()
df_molde['RA'] = df_ativos['RA']
df_molde['Nome'] = df_ativos['Nome_Limpo']
df_molde['Telefone'] = df_ativos['Telefone_Limpo']
df_molde['E-mail'] = df_ativos['Email_Limpo']
df_molde['country code'] = '55'

# Mapeando a Filial para a coluna Tags
df_molde['Tags'] = df_ativos['FILIAL'].astype(str).str.strip()
df_molde['Tags'] = df_molde['Tags'].replace(['nan', 'None'], '')

# ==========================================
# 6. DESDUPLICAÇÃO FINAL
# ==========================================
df_com_ra = df_molde[df_molde['RA'].notna()]
df_sem_ra = df_molde[df_molde['RA'].isna()]

df_com_ra = df_com_ra.drop_duplicates(subset=['RA'], keep='first')
df_molde = pd.concat([df_com_ra, df_sem_ra], ignore_index=True)
df_molde = df_molde.drop_duplicates(subset=['Telefone'], keep='first')

metricas["Duplicados"] = qtd_atual - len(df_molde)
qtd_atual = len(df_molde)

# ==========================================
# 7. EXPORTAÇÃO E RELATÓRIO
# ==========================================
colunas_finais = ['Nome', 'Telefone', 'country code', 'Tags', 'E-mail']
df_final = df_molde[colunas_finais]

arquivo_saida = os.path.join(caminho_output, "BaseAlunosAtivos_Formatada.xlsx")
df_final.to_excel(arquivo_saida, index=False)

print("\n" + "="*60)
print(f"✅ SUCESSO! Base de Alunos Ativos formatada no molde das Filiais.")
print(f"📁 Salvo em: {arquivo_saida}")
print("="*60)

print("\n📊 RELATÓRIO DE EXCLUSÕES (O FUNIL):")
print("-" * 60)
print(f"📉 Total Original ({' + '.join(ABAS_ENTRADA)}): {qtd_original} alunos")
print(f"❌ Nomes Inválidos: -{metricas['Nomes_Invalidos']}")
print(f"❌ Sem Telefone Válido (aluno e responsáveis): -{metricas['Sem_Telefone']}")
print(f"👯 Cortados por Duplicação (RA/Telefone): -{metricas['Duplicados']}")
print("-" * 60)
print(f"🚀 TOTAL FINAL PRONTO PARA DISPARO: {len(df_final)} alunos")
print("="*60)

print("\n🏢 RESUMO DE ALUNOS POR FILIAL (TAGS):")
print("-" * 60)
contagem_tags = df_final['Tags'].value_counts()
for filial, quantidade in contagem_tags.items():
    if filial != '':
        print(f"🔹 {filial}: {quantidade} aluno(s)")
    else:
        print(f"🔸 (Sem Filial Cadastrada): {quantidade} aluno(s)")
print("-" * 60)
