import pandas as pd
import re
import os
import sys
import unicodedata
import warnings
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

warnings.filterwarnings('ignore')

# ==========================================
# 1. DEFINIÇÃO DE CAMINHOS FIXOS
# ==========================================
# Uso: python formatar_formandos.py [planilha_entrada.xlsx] [pasta_saida] [depara_turmas.xlsx]
arquivo_entrada = r"C:\Users\everton.araujo\Desktop\Disparo_Limpeza\Base_Ativos\Formandos_2026.2.xlsx"
caminho_output = r"C:\Users\everton.araujo\Desktop\Disparo_Limpeza"
# Export do TOTVS do mesmo semestre com a coluna TURMA (opcional)
arquivo_depara = ''

if len(sys.argv) > 1: arquivo_entrada = sys.argv[1]
if len(sys.argv) > 2: caminho_output = sys.argv[2]
if len(sys.argv) > 3: arquivo_depara = sys.argv[3]

if not os.path.exists(caminho_output):
    os.makedirs(caminho_output)

NOME_SAIDA = 'Formandos_2026.2.xlsx'

# Níveis de quem forma no semestre -> grupo exibido na planilha
NIVEIS_FORMANDOS = {
    'A4': 'Advanced (A4)',
    'MF2': 'Masterflex (MF2)',
    'IFMF2': 'IBEU Fit Masterflex (IFMF2)',
}

# Alunos citados nominalmente (ex.: formandos do High School)
ALUNOS_EXTRAS = {
    'CELINA MOSSÉ MUANIS': 'High School',
    'CLARA TAVARES PAES MUANIS': 'High School',
    'SOFIA ANDRADE VALLIN': 'High School',
}
ORDEM_GRUPOS = list(NIVEIS_FORMANDOS.values()) + ['High School']

RENOMEAR_FILIAIS = {
    'INSTITUTO BRASIL ESTADOS UNIDOS': 'IBEU AT HOME',
}

# ==========================================
# 2. FUNÇÕES DE LIMPEZA
# ==========================================
def limpar_nome(nome):
    if pd.isna(nome) or str(nome).strip() == '':
        return ''
    nome_str = str(nome).strip().upper()
    nome_str = re.sub(r'\s*\(\d+\)\s*', '', nome_str)
    nome_str = re.sub(r'[^A-ZÁÀÂÃÉÈÊÍÏÓÔÕÖÚÇÑ\s\']', '', nome_str)
    return re.sub(r'\s+', ' ', nome_str).strip()

def nome_comparacao(nome):
    nome_str = unicodedata.normalize('NFKD', limpar_nome(nome)).encode('ascii', 'ignore').decode()
    return re.sub(r'\s+', ' ', nome_str).strip()

def limpar_email(email):
    if pd.isna(email):
        return ''
    # Uma célula pode ter vários e-mails ("a@x.com;b@y.com", "p fulano@gmail.com")
    encontrados = re.findall(r'[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}', str(email))
    return juntar([e.lower() for e in encontrados])

def nome_pessoa(nome):
    nome_str = limpar_nome(nome)
    return '' if nome_str in ['SEM RESP', 'NAN'] else nome_str

def juntar(valores):
    vistos = []
    for v in valores:
        if v and v not in vistos:
            vistos.append(v)
    return '; '.join(vistos)

# ==========================================
# 3. LEITURA (EXPORT EM TABELA DINÂMICA DO TOTVS)
# ==========================================
def ler_totvs(arquivo):
    bruto = pd.read_excel(arquivo, header=None, dtype=str)

    # O cabeçalho real é a linha que tem a célula "FILIAL"
    linha_cabecalho = bruto.index[bruto.apply(lambda l: (l.astype(str).str.strip() == 'FILIAL').any(), axis=1)][0]
    df = pd.read_excel(arquivo, header=linha_cabecalho, dtype=str).dropna(axis=1, how='all')
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, ~df.columns.str.startswith('Unnamed')]
    df = df[df['FILIAL'].astype(str).str.strip().str.upper() != 'TOTAL GERAL']
    # Linhas de subtotal da tabela dinâmica ("ALUNO ATIVO Total") e linhas em branco
    df = df[~df['STATUS'].astype(str).str.upper().str.contains('TOTAL')]

    # Células mescladas: a linha em branco repete a filial/aluno da linha de cima
    df['FILIAL'] = df['FILIAL'].ffill().str.strip()
    df['NOME'] = df['NOME'].ffill().str.strip()
    df['FILIAL'] = df['FILIAL'].replace(RENOMEAR_FILIAIS)
    df = df[df['NIVEL'].notna()]
    df['NIVEL'] = df['NIVEL'].astype(str).str.strip().str.upper()
    df['RA'] = df['NOME'].str.extract(r'\((\d+)\)')[0]
    df['Nome_Comparacao'] = df['NOME'].apply(nome_comparacao)
    return df

print("Lendo base de formandos (TOTVS)...")
try:
    df = ler_totvs(arquivo_entrada)
    df_depara = ler_totvs(arquivo_depara) if arquivo_depara else None
except FileNotFoundError as e:
    print(f"❌ Erro ao ler a base: {e}")
    exit()

def buscar_turma(ra, niveis, filial):
    """Turma do aluno no DEPARA; prioriza matrícula ativa na mesma filial (ignora transferências)."""
    if df_depara is None:
        return '', ''
    cand = df_depara[(df_depara['RA'] == ra) & (df_depara['NIVEL'].isin(niveis))]
    if cand.empty:
        return 'NÃO ENCONTRADA', ''
    ativo = cand['STATUS'].str.contains('ATIVO', na=False)
    mesma_filial = cand['FILIAL'] == filial
    cand = cand.assign(prioridade=ativo.astype(int) * 2 + mesma_filial.astype(int)).sort_values('prioridade', ascending=False)
    melhor = cand.iloc[0]
    return str(melhor['TURMA']).strip(), str(melhor['STATUS']).strip()

# ==========================================
# 4. SELEÇÃO DOS FORMANDOS
# ==========================================
df_niveis = df[df['NIVEL'].isin(NIVEIS_FORMANDOS)].copy()
df_niveis['Grupo'] = df_niveis['NIVEL'].map(NIVEIS_FORMANDOS)

extras_encontrados, extras_nao_encontrados = [], []
partes_extras = []
for nome, grupo in ALUNOS_EXTRAS.items():
    achado = df[df['Nome_Comparacao'] == nome_comparacao(nome)].copy()
    if len(achado):
        achado['Grupo'] = grupo
        partes_extras.append(achado)
        extras_encontrados.append(nome)
    else:
        extras_nao_encontrados.append(nome)

df_sel = pd.concat([df_niveis] + partes_extras, ignore_index=True)

# ==========================================
# 5. UMA LINHA POR ALUNO, COM TODOS OS E-MAILS
# ==========================================
registros = []
for chave, grupo_linhas in df_sel.groupby(df_sel['RA'].fillna(df_sel['Nome_Comparacao']), sort=False):
    linha = grupo_linhas.iloc[0]
    email_aluno = juntar(grupo_linhas['EMAIL'].apply(limpar_email))
    email_fin = juntar(grupo_linhas['EMAIL RESP FIN'].apply(limpar_email))
    email_aca = juntar(grupo_linhas['EMAIL RESP ACA'].apply(limpar_email))
    registros.append({
        'Filial': linha['FILIAL'],
        'Grupo': linha['Grupo'],
        'Curso': juntar(grupo_linhas['CURSO'].str.strip()),
        'Nível': juntar(grupo_linhas['NIVEL']),
        'Turma': '',
        'Situação na Turma': '',
        'RA': linha['RA'] if pd.notna(linha['RA']) else '',
        'Aluno': limpar_nome(linha['NOME']),
        'Faixa Etária': str(linha['FAIXA ETÁRIA']).strip(),
        'Status': juntar(grupo_linhas['STATUS'].str.strip()),
        'E-mail do Aluno': email_aluno,
        'Responsável Financeiro': nome_pessoa(linha['NOME RESP FIN']),
        'E-mail Resp. Financeiro': email_fin,
        'Responsável Acadêmico': nome_pessoa(linha['NOME RESP ACA']),
        'E-mail Resp. Acadêmico': email_aca,
        'Todos os E-mails': juntar('; '.join([email_aluno, email_fin, email_aca]).split('; ')),
    })

df_final = pd.DataFrame(registros)
if df_depara is not None:
    turmas = df_final.apply(lambda l: buscar_turma(l['RA'], l['Nível'].split('; '), l['Filial']), axis=1)
    df_final['Turma'] = [t for t, _ in turmas]
    df_final['Situação na Turma'] = [st for _, st in turmas]
else:
    df_final = df_final.drop(columns=['Turma', 'Situação na Turma'])
df_final['Grupo'] = pd.Categorical(df_final['Grupo'], ORDEM_GRUPOS, ordered=True)
df_final = df_final.sort_values(['Grupo', 'Filial', 'Aluno']).reset_index(drop=True)
df_final['Grupo'] = df_final['Grupo'].astype(str)

emails_unicos = juntar('; '.join(df_final['Todos os E-mails']).split('; ')).split('; ')
emails_unicos = [e for e in emails_unicos if e]
df_emails = pd.DataFrame({'E-mail': emails_unicos})

df_nao_encontrados = pd.DataFrame({
    'Aluno': extras_nao_encontrados,
    'Grupo': [ALUNOS_EXTRAS[n] for n in extras_nao_encontrados],
    'Observação': 'Não encontrado(a) na base do TOTVS',
})

# ==========================================
# 6. EXPORTAÇÃO
# ==========================================
COR_CABECALHO = '1F4E78'
COR_ZEBRA = 'EAF1F8'
COR_SEM_EMAIL = 'FCE4D6'
borda = Border(bottom=Side(style='thin', color='BFBFBF'))

def formatar_aba(ws, df_aba, sem_email=None):
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = ws.dimensions
    ws.sheet_view.showGridLines = False

    for cel in ws[1]:
        cel.font = Font(bold=True, color='FFFFFF')
        cel.fill = PatternFill('solid', fgColor=COR_CABECALHO)
        cel.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws.row_dimensions[1].height = 30

    for i, linha in enumerate(ws.iter_rows(min_row=2), start=0):
        destacar = sem_email is not None and sem_email[i]
        cor = COR_SEM_EMAIL if destacar else (COR_ZEBRA if i % 2 else None)
        for cel in linha:
            cel.border = borda
            cel.alignment = Alignment(vertical='center')
            if cor: cel.fill = PatternFill('solid', fgColor=cor)

    for idx, coluna in enumerate(df_aba.columns, start=1):
        letra = get_column_letter(idx)
        if coluna == 'RA':
            for cel in ws[letra][1:]:
                cel.number_format = '@'
        maior = max([len(str(coluna))] + [len(str(v)) for v in df_aba[coluna].fillna('')])
        ws.column_dimensions[letra].width = min(max(maior + 3, 10), 60)

# Colunas de uso interno que não vão para a planilha
COLUNAS_OCULTAS = ['Todos os E-mails']

def escrever_aba(writer, nome_aba, df_aba):
    sem_email = (df_aba['Todos os E-mails'] == '').tolist() if 'Todos os E-mails' in df_aba.columns else None
    df_aba = df_aba.drop(columns=[c for c in COLUNAS_OCULTAS if c in df_aba.columns])
    df_aba.to_excel(writer, sheet_name=nome_aba, index=False)
    formatar_aba(writer.sheets[nome_aba], df_aba, sem_email)

arquivo_saida = os.path.join(caminho_output, NOME_SAIDA)
with pd.ExcelWriter(arquivo_saida, engine='openpyxl') as writer:
    escrever_aba(writer, 'Todos', df_final)
    for grupo in ORDEM_GRUPOS:
        df_grupo = df_final[df_final['Grupo'] == grupo]
        if len(df_grupo):
            escrever_aba(writer, grupo[:31], df_grupo)
    escrever_aba(writer, 'Lista de E-mails', df_emails)
    if len(df_nao_encontrados):
        escrever_aba(writer, 'Não Encontrados', df_nao_encontrados)

# ==========================================
# 7. RELATÓRIO
# ==========================================
print("\n" + "="*60)
print(f"✅ SUCESSO! Base de formandos gerada.")
print(f"📁 Salvo em: {arquivo_saida}")
print("="*60)

print("\n📊 RELATÓRIO:")
print("-" * 60)
print(f"📉 Linhas na base do TOTVS: {len(df)}")
for grupo in ORDEM_GRUPOS:
    print(f"🎓 {grupo}: {(df_final['Grupo'] == grupo).sum()} aluno(s)")
print(f"✅ Extras encontrados: {', '.join(extras_encontrados) or '-'}")
print(f"❌ Extras NÃO encontrados: {', '.join(extras_nao_encontrados) or '-'}")
if df_depara is not None:
    nao_ativos = df_final[~df_final['Situação na Turma'].str.contains('ATIVO')]
    print(f"🏫 Turma encontrada: {(df_final['Turma'] != 'NÃO ENCONTRADA').sum()} de {len(df_final)}")
    for _, l in nao_ativos.iterrows():
        print(f"     ⚠️ {l['Aluno']} ({l['Filial']}) — turma {l['Turma']}: {l['Situação na Turma'] or 'sem turma'}")
print(f"🎗️ Bolsistas: {df_final['Status'].str.contains('BOLSISTA').sum()}")
print(f"📧 Alunos sem nenhum e-mail: {(df_final['Todos os E-mails'] == '').sum()}")
print("-" * 60)
print(f"🚀 TOTAL DE FORMANDOS: {len(df_final)} | E-MAILS ÚNICOS: {len(df_emails)}")
print("="*60)
