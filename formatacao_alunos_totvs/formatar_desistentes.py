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
# Uso: python formatar_desistentes.py [desistentes.xlsx] [pasta_saida] [base_ativos1.xlsx base_ativos2.xlsx ...]
arquivo_desistentes = r"C:\Users\everton.araujo\Desktop\Disparo_Limpeza\Desistentes_2024_e_2025.xlsx"
caminho_output = r"C:\Users\everton.araujo\Desktop\Disparo_Limpeza"
arquivos_ativos = [r"C:\Users\everton.araujo\Desktop\Disparo_Limpeza\Base_Ativos\lista_alunos.xlsx"]

if len(sys.argv) > 1: arquivo_desistentes = sys.argv[1]
if len(sys.argv) > 2: caminho_output = sys.argv[2]
if len(sys.argv) > 3: arquivos_ativos = sys.argv[3:]

if not os.path.exists(caminho_output):
    os.makedirs(caminho_output)

NOME_SAIDA = 'Desistentes_2024_2025_por_Filial.xlsx'

# Cursos regulares: os da aba REGULARES da base de ativos + infantis confirmados
CURSOS_REGULARES = [
    'ADVANCED', 'BASIC', 'BASIC FIT', 'BEST', 'BEST CONVERSATION 4', 'BEST PLUS',
    'CONVERSATION', 'CONVERSATION 120', 'CONVERSATION FIT', 'CONVERSATION ONCE',
    'IBEU FIT FLEX SEMESTRAL', 'IBEU FIT MASTER FLEX', 'IBEU FIT TOP FLEX SEMESTRAL',
    'IBEU FLEX SEMESTRAL', 'INTERMEDIATE', 'KIDS EXPEDITION', 'KIDS FRONTIER',
    'KIDS JOURNEY', 'MASTER FLEX', 'POP KIDS', 'TEENS', 'TOP FLEX SEMESTRAL',
    'MAGIC CLUB', 'KIDS', 'KIDS ADVENTURE',
]

# Nas bases de ativos, linhas com estes status não contam como aluno ativo
STATUS_NAO_ATIVOS = r'DESISTENTE|CANCELAD|TRANSFERIDO|TRANCADO'

# Grupos de contato: cada um tem nome, e-mail e telefones
CONTATOS = {
    'Resp. Fin.': {
        'nome': 'NOME RESP FIN',
        'email': 'EMAIL RESP FIN',
        'telefones': ['TELEFONE 1 RESP. FIN.', 'TELEFONE 2 RESP. FIN.', 'TELEFONE COMERCIAL RESP. FIN.'],
    },
    'Aluno': {
        'nome': 'NOME',
        'email': 'EMAIL',
        'telefones': ['TELEFONE 1 ALUNO', 'TELEFONE 2 ALUNO', 'TELEFONE 3 ALUNO'],
    },
    'Resp. Acad.': {
        'nome': 'NOME RESP ACA',
        'email': 'EMAIL RESP ACA',
        'telefones': ['TELEFONE 1 RESP. ACAD.', 'TELEFONE 2 RESP. ACAD.', 'TELEFONE 3 RESP. ACAD.'],
    },
}

# Ordem de busca do contato (só celular; fixo nunca entra)
ORDEM_MENOR = ['Resp. Fin.', 'Resp. Acad.']
ORDEM_ADULTO = ['Resp. Fin.', 'Aluno', 'Resp. Acad.']
# Adulto cujo pagante é empresa: usa o contato do próprio aluno (nunca o da empresa)
ORDEM_ADULTO_EMPRESA = ['Aluno', 'Resp. Acad.']

# Pagante pessoa jurídica (empresa/escola que paga o curso)
PADRAO_EMPRESA = r'\bLTDA\b|\bS\.?\s?A\.?$|S/A|EIRELI|ESCOLA|CORRETORA|COMPANY|EMPRESA|CONSULTING|CONSULTORIA'

RENOMEAR_FILIAIS = {
    'INSTITUTO BRASIL ESTADOS UNIDOS': 'IBEU AT HOME',
}

# ==========================================
# 2. FUNÇÕES DE LIMPEZA
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
    nome_str = re.sub(r'\s*\(\d+\)\s*', '', nome_str)
    nome_str = re.sub(r'[^A-ZÁÀÂÃÉÈÊÍÏÓÔÕÖÚÇÑ\s\']', '', nome_str)
    return re.sub(r'\s+', ' ', nome_str).strip()

def nome_comparacao(nome):
    nome_str = unicodedata.normalize('NFKD', limpar_nome(nome)).encode('ascii', 'ignore').decode()
    return re.sub(r'\s+', ' ', nome_str).strip()

def limpar_email(email):
    if pd.isna(email):
        return ''
    # Uma célula pode ter vários e-mails; fica o primeiro válido
    encontrados = re.findall(r'[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}', str(email))
    return encontrados[0].lower() if encontrados else ''

def is_celular(numero):
    # 55 + DDD + 9 dígitos começando com 9
    return len(numero) == 13 and numero[4] == '9'

def is_menor(faixa):
    # "13 - 17 ANOS" -> menor | "18 - 25 ANOS" / "61 ou + ANOS" -> adulto | "05 ou - ANOS" -> menor
    faixa_str = str(faixa).upper()
    numeros = [int(n) for n in re.findall(r'\d+', faixa_str)]
    if not numeros: return False
    if 'OU +' in faixa_str: return False
    return max(numeros) < 18

def buscar_contato(linha, ordem):
    """Busca celular e e-mail separadamente, seguindo a ordem dos grupos."""
    celular, origem_cel, email = '', '', ''
    for grupo in ordem:
        cfg = CONTATOS[grupo]
        if not celular:
            for col in cfg['telefones']:
                tel = formatar_telefone(linha.get(col))
                if tel and is_celular(tel):
                    celular, origem_cel = tel, grupo
                    break
        if not email:
            email = limpar_email(linha.get(cfg['email']))
    return celular, origem_cel, email

def nome_pessoa(nome):
    nome_str = limpar_nome(nome)
    return '' if nome_str in ['SEM RESP', 'NAN'] else nome_str

def juntar(valores):
    vistos = []
    for v in valores:
        if v and v not in vistos:
            vistos.append(v)
    return '; '.join(vistos)

def ler_totvs(arquivo):
    """Lê todas as abas de um export do TOTVS (simples ou em tabela dinâmica)."""
    partes = []
    for aba, bruto in pd.read_excel(arquivo, sheet_name=None, header=None, dtype=str).items():
        achou = bruto.apply(lambda l: (l.astype(str).str.strip() == 'FILIAL').any(), axis=1)
        if not achou.any():
            continue
        df = pd.read_excel(arquivo, sheet_name=aba, header=bruto.index[achou][0], dtype=str).dropna(axis=1, how='all')
        df.columns = [str(c).strip() for c in df.columns]
        df = df.loc[:, ~df.columns.str.startswith('Unnamed')]
        # Linhas de subtotal da tabela dinâmica ("ALUNO DESISTENTE Total")
        df = df[~df['STATUS'].astype(str).str.upper().str.contains('TOTAL')]
        # Células mescladas: a linha em branco repete a filial/aluno da linha de cima
        df['FILIAL'] = df['FILIAL'].ffill().str.strip().replace(RENOMEAR_FILIAIS)
        df['NOME'] = df['NOME'].ffill().str.strip()
        df = df[df['CURSO'].notna()]
        df['RA'] = df['NOME'].str.extract(r'\((\d+)\)')[0]
        df['Nome_Comparacao'] = df['NOME'].apply(nome_comparacao)
        partes.append(df)
    return pd.concat(partes, ignore_index=True)

# ==========================================
# 3. LEITURA DAS BASES
# ==========================================
print("Lendo base de desistentes e bases de ativos (TOTVS)...")
try:
    df = ler_totvs(arquivo_desistentes)
    ativos = [ler_totvs(arquivo) for arquivo in arquivos_ativos]
except FileNotFoundError as e:
    print(f"❌ Erro ao ler arquivos: {e}")
    exit()

df_ativos = pd.concat(ativos, ignore_index=True)
df_ativos = df_ativos[~df_ativos['STATUS'].astype(str).str.upper().str.contains(STATUS_NAO_ATIVOS)]
ras_ativos = set(df_ativos['RA'].dropna())
nomes_ativos = set(df_ativos['Nome_Comparacao']) - {''}

df['CURSO'] = df['CURSO'].str.strip()
df['Regular'] = df['CURSO'].str.upper().isin(CURSOS_REGULARES)
df['Ativo'] = df['RA'].isin(ras_ativos) | df['Nome_Comparacao'].isin(nomes_ativos)

# ==========================================
# 4. CONTATO DE CADA LINHA
# ==========================================
contatos = []
for _, linha in df.iterrows():
    menor = is_menor(linha.get('FAIXA ETÁRIA'))
    pagante_empresa = bool(re.search(PADRAO_EMPRESA, str(linha.get('NOME RESP FIN', '')).upper()))
    if menor: ordem = ORDEM_MENOR
    elif pagante_empresa: ordem = ORDEM_ADULTO_EMPRESA
    else: ordem = ORDEM_ADULTO
    contatos.append((menor,) + buscar_contato(linha, ordem))
df['Menor'] = [c[0] for c in contatos]
df['Celular'] = [c[1] for c in contatos]
df['Origem do Celular'] = [c[2] for c in contatos]
df['E-mail'] = [c[3] for c in contatos]

# ==========================================
# 5. FUNIL POR FILIAL
# ==========================================
resumo = {}
def contar(filial, chave, qtd=1):
    resumo.setdefault(filial, {}).setdefault(chave, 0)
    resumo[filial][chave] += qtd

for filial, qtd in df.groupby('FILIAL')['RA'].nunique().items():
    contar(filial, 'Total no Cubo', qtd)

df_reg = df[df['Regular']]
for filial, grupo in df.groupby('FILIAL'):
    contar(filial, 'Cursos Não Regulares', grupo['RA'].nunique() - grupo[grupo['Regular']]['RA'].nunique())

df_livre = df_reg[~df_reg['Ativo']]
for filial, grupo in df_reg.groupby('FILIAL'):
    contar(filial, 'Já Ativos', grupo['RA'].nunique() - grupo[~grupo['Ativo']]['RA'].nunique())

# Um aluno por RA: junta cursos/status e usa o primeiro celular encontrado
registros = []
for (filial, ra), grupo in df_livre.groupby(['FILIAL', df_livre['RA'].fillna(df_livre['Nome_Comparacao'])], sort=False):
    com_cel = grupo[grupo['Celular'] != '']
    linha = com_cel.iloc[0] if len(com_cel) else grupo.iloc[0]
    registros.append({
        'Filial': filial,
        'Aluno': limpar_nome(linha['NOME']),
        'RA': linha['RA'] if pd.notna(linha['RA']) else '',
        'Curso': juntar(grupo['CURSO']),
        'Status': juntar(grupo['STATUS'].str.strip()),
        'Tipo': juntar(grupo['TIPOAL'].fillna('').str.strip()),
        'Faixa Etária': str(linha['FAIXA ETÁRIA']).strip(),
        'Idade': 'Menor' if linha['Menor'] else 'Adulto',
        'Responsável Financeiro': nome_pessoa(linha['NOME RESP FIN']),
        'Celular': linha['Celular'],
        'E-mail': juntar(grupo['E-mail']).split('; ')[0],
        'Origem do Contato': linha['Origem do Celular'],
    })
df_final = pd.DataFrame(registros)

for filial, grupo in df_final.groupby('Filial'):
    contar(filial, 'Sem Celular', (grupo['Celular'] == '').sum())

df_final = df_final[df_final['Celular'] != '']
df_final = df_final.sort_values(['Filial', 'Status', 'Aluno']).reset_index(drop=True)

linhas_resumo = []
for filial in sorted(resumo):
    r = resumo[filial]
    disp = df_final[df_final['Filial'] == filial]
    linhas_resumo.append({
        'Filial': filial,
        'Total no Cubo': r.get('Total no Cubo', 0),
        'Cursos Não Regulares': -r.get('Cursos Não Regulares', 0),
        'Já Ativos': -r.get('Já Ativos', 0),
        'Sem Celular': -r.get('Sem Celular', 0),
        'Disponível para Trabalho': len(disp),
        'Desistentes': disp['Status'].str.contains('DESISTENTE').sum(),
        # Aluno com os dois status (um curso de cada) conta como desistente
        'Reserva Cancelada': (~disp['Status'].str.contains('DESISTENTE')).sum(),
    })
df_resumo = pd.DataFrame(linhas_resumo).sort_values('Disponível para Trabalho', ascending=False)
total = df_resumo.drop(columns='Filial').sum()
total['Filial'] = 'TOTAL'
df_resumo = pd.concat([df_resumo, pd.DataFrame([total])], ignore_index=True)

# ==========================================
# 6. EXPORTAÇÃO
# ==========================================
COR_CABECALHO = '1F4E78'
COR_ZEBRA = 'EAF1F8'
borda = Border(bottom=Side(style='thin', color='BFBFBF'))

def formatar_aba(ws, df_aba, linha_total=False):
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = ws.dimensions
    ws.sheet_view.showGridLines = False

    for cel in ws[1]:
        cel.font = Font(bold=True, color='FFFFFF')
        cel.fill = PatternFill('solid', fgColor=COR_CABECALHO)
        cel.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws.row_dimensions[1].height = 30

    for i, linha in enumerate(ws.iter_rows(min_row=2), start=0):
        for cel in linha:
            cel.border = borda
            cel.alignment = Alignment(vertical='center')
            if i % 2: cel.fill = PatternFill('solid', fgColor=COR_ZEBRA)
    if linha_total:
        for cel in ws[ws.max_row]:
            cel.font = Font(bold=True)
            cel.border = Border(top=Side(style='medium', color=COR_CABECALHO))

    for idx, coluna in enumerate(df_aba.columns, start=1):
        letra = get_column_letter(idx)
        if coluna in ['Celular', 'RA']:
            for cel in ws[letra][1:]:
                cel.number_format = '@'
        maior = max([len(str(coluna))] + [len(str(v)) for v in df_aba[coluna].fillna('')])
        ws.column_dimensions[letra].width = min(max(maior + 3, 10), 50)

arquivo_saida = os.path.join(caminho_output, NOME_SAIDA)
with pd.ExcelWriter(arquivo_saida, engine='openpyxl') as writer:
    df_resumo.to_excel(writer, sheet_name='Resumo', index=False)
    formatar_aba(writer.sheets['Resumo'], df_resumo, linha_total=True)
    # Uma aba por filial (sem aba geral)
    for filial in df_resumo['Filial']:
        df_filial = df_final[df_final['Filial'] == filial].drop(columns='Filial')
        if len(df_filial):
            nome_aba = re.sub(r'[\[\]\*\?/\\:]', '', filial)[:31]
            df_filial.to_excel(writer, sheet_name=nome_aba, index=False)
            formatar_aba(writer.sheets[nome_aba], df_filial)

# ==========================================
# 7. RELATÓRIO
# ==========================================
t = df_resumo.iloc[-1]
print("\n" + "="*60)
print(f"✅ SUCESSO! Base de desistentes separada por filial.")
print(f"📁 Salvo em: {arquivo_saida}")
print("="*60)
print("\n📊 RELATÓRIO DE EXCLUSÕES (O FUNIL):")
print("-" * 60)
print(f"📉 Alunos no cubo: {t['Total no Cubo']}")
print(f"📚 Cursos não regulares: {t['Cursos Não Regulares']}")
print(f"🎓 Já estão ativos (RA ou nome nas bases de ativos): {t['Já Ativos']}")
print(f"☎️ Sem celular: {t['Sem Celular']}")
print("-" * 60)
print(f"🚀 DISPONÍVEL PARA TRABALHO: {t['Disponível para Trabalho']} "
      f"({t['Desistentes']} desistentes / {t['Reserva Cancelada']} reserva cancelada)")
print("="*60)
print("\n🏢 POR FILIAL:")
for _, r in df_resumo.iloc[:-1].iterrows():
    print(f"🔹 {r['Filial']}: {r['Disponível para Trabalho']}")
