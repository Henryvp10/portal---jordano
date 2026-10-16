import os
import re
import unicodedata
import pythoncom
import win32com.client as win32
import smtplib
import calendar
import pandas as pd
from datetime import datetime, timedelta
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from flask import Flask, request, jsonify
from flask_cors import CORS
from pypdf import PdfReader

app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "*"}}, supports_credentials=True)

# Configuração de credenciais SMTP
EMAIL_USER = "henry.perissinotti@transjordano.com.br"
EMAIL_PASS = "henry2005@" 

# Caminhos corporativos da rede
PASTA_MOTORISTAS = r"G:\Drives compartilhados\SSMA\SSMA\Arquivos da Qualidade - Seguranca\3. Motoristas\1. Motoristas"
PASTA_VEICULOS = r"\\192.168.0.105\Geral\MANUTENÇÃO\DOCUMENTAÇÃO\FROTA TRANSJORDANO"
PASTA_LICENCAS = r"\\192.168.0.105\Geral\MANUTENÇÃO\DOCUMENTAÇÃO\FROTA TRANSJORDANO\LICENÇAS"

# Caminho apontando diretamente para static/Colaboradores.xlsx
ARQUIVO_PLANILHA = os.path.join(os.path.dirname(__file__), 'static', 'Colaboradores.xlsx')

MAPEAMENTO_MOTORISTAS = {
    'cnh': ['cnh', 'habilitacao', 'carteira', 'driver'],
    'aso': ['aso', 'atestado', 'saude', 'medico'],
    'mopp': ['mopp', 'curso mopp'],
    'nr20': ['nr20', 'nr-20', 'nr 20'],
    'nr35': ['nr35', 'nr-35', 'nr 35'],
    'toxicologico': ['toxicologico', 'toxico', 'toxic'],
    'psicologica': ['psicologica', 'av.psicologica', 'avaliacao', 'avaliacaopsicologica', 'psico', 'ava psico']
}

MAPEAMENTO_PASTAS_VEICULOS = {
    'calibragem': ['CALIBRAGEM'],
    'cipp': ['CIPP'],
    'civ': ['CIV - CAVALOS', 'CIV - CARRETAS'],
    'crlv': ['CRLV'],
    'cronotacografo': ['CRONOTACOGRAFO']
}

MAPEAMENTO_PASTAS_LICENCAS = {
    'AET': 'AETs',
    'IBAMA_AMBIENTAL': 'Ambiental IBAMA - Licença ambiental e Certificado de Regularidade',
    'IBAMA / AMBIENTAL': 'Ambiental IBAMA - Licença ambiental e Certificado de Regularidade',
    'CADASTRO_TECNICO_FEDERAL': 'Ambiental IBAMA - Licença ambiental e Certificado de Regularidade',
    'CADASTRO FEDERAL': 'Ambiental IBAMA - Licença ambiental e Certificado de Regularidade',
    'LETPP': 'LETPP - DSV',
    'LICENCA_OPERACAO': 'LICENÇAS DE OPERAÇÃO',
    'LICENÇA DE OPERAÇÃO': 'LICENÇAS DE OPERAÇÃO',
    'ANTT': ''
}


def remover_acentos(texto):
    if not texto:
        return ""
    texto_normalizado = unicodedata.normalize('NFD', str(texto))
    return "".join(c for c in texto_normalizado if unicodedata.category(c) != 'Mn')


def limpar_string(texto):
    """Remove tudo o que não for número."""
    return re.sub(r'\D', '', str(texto)) if texto else ""


# --- FUNÇÕES AUXILIARES DE VENCIMENTO DO CRLV ---
def extrair_ultimo_digito_placa(nome_arquivo):
    """Localiza o padrão de placa no nome do arquivo e retorna o último dígito numérico."""
    match = re.search(r'([A-Z]{3}\d[A-Z0-9]\d{2})', nome_arquivo.upper())
    if match:
        placa = match.group(1)
        numeros = re.findall(r'\d', placa)
        if numeros:
            return int(numeros[-1])
    
    numeros_gerais = re.findall(r'\d', nome_arquivo)
    if numeros_gerais:
        return int(numeros_gerais[-1])
    return None


def extrair_peso_bruto_pdf(caminho_arquivo):
    """Abre o PDF do CRLV e lê o campo Peso Bruto Total para diferenciar Cavalo de Carreta."""
    try:
        reader = PdfReader(caminho_arquivo)
        texto_completo = ""
        for page in reader.pages:
            texto_completo += page.extract_text() or ""
        
        texto_norm = remover_acentos(texto_completo.upper())

        if '23.0' in texto_norm or '23,0' in texto_norm or '23000' in texto_norm:
            return 'CAVALO'
        elif '25.5' in texto_norm or '25,5' in texto_norm or '25500' in texto_norm:
            return 'CARRETA'
    except Exception as e:
        print(f"[ERRO LEITURA PDF CRLV] Não foi possível ler {caminho_arquivo}: {e}")
    
    return None


def calcular_vencimento_crlv(nome_arquivo, caminho_completo=None):
    """Calcula a data exata de vencimento do CRLV baseado no ano do exercício, final da placa e PBT."""
    nome_upper = nome_arquivo.upper()
    
    # Busca segura do ano de 4 dígitos no nome do arquivo
    ano_match = re.search(r'CRLV[_\s-]*(\d{4})', nome_upper)
    if not ano_match:
        ano_match = re.search(r'(20\d{2})', nome_upper)
        
    if not ano_match or not ano_match.groups():
        return None

    try:
        ano_exercicio = int(ano_match.group(1))
    except (IndexError, ValueError):
        return None

    ano_vencimento = ano_exercicio + 1

    ultimo_digito = extrair_ultimo_digito_placa(nome_upper)
    if ultimo_digito is None:
        return None

    is_cavalo = None
    if '_CV' in nome_upper or 'CAVALO' in nome_upper:
        is_cavalo = True
    elif '_SR' in nome_upper or 'CARRETA' in nome_upper or 'REBOQUE' in nome_upper:
        is_cavalo = False

    if is_cavalo is None and caminho_completo and os.path.exists(caminho_completo):
        tipo_pdf = extrair_peso_bruto_pdf(caminho_completo)
        if tipo_pdf == 'CAVALO':
            is_cavalo = True
        elif tipo_pdf == 'CARRETA':
            is_cavalo = False

    if is_cavalo is None:
        is_cavalo = True

    mes_vencimento = None

    if is_cavalo:
        if ultimo_digito in [1, 2]:
            mes_vencimento = 9
        elif ultimo_digito in [3, 4, 5]:
            mes_vencimento = 10
        elif ultimo_digito in [6, 7, 8]:
            mes_vencimento = 11
        elif ultimo_digito in [9, 0]:
            mes_vencimento = 12
    else:
        if ultimo_digito in [1, 2]:
            mes_vencimento = 7
        elif ultimo_digito in [3, 4]:
            mes_vencimento = 8
        elif ultimo_digito in [5, 6]:
            mes_vencimento = 9
        elif ultimo_digito in [7, 8]:
            mes_vencimento = 10
        elif ultimo_digito == 9:
            mes_vencimento = 11
        elif ultimo_digito == 0:
            mes_vencimento = 12

    if mes_vencimento:
        ultimo_dia = calendar.monthrange(ano_vencimento, mes_vencimento)[1]
        return datetime(ano_vencimento, mes_vencimento, ultimo_dia)

    return None


def extrair_data_do_nome(nome_arquivo, caminho_completo=None):
    """Procura por padrões de data explícitos ou calcula a regra especial do CRLV."""
    if 'CRLV' in nome_arquivo.upper():
        dt_crlv = calcular_vencimento_crlv(nome_arquivo, caminho_completo)
        if dt_crlv:
            return dt_crlv

    padrao_br = re.search(r'(\d{2})[-._](\d{2})[-._](\d{4})', nome_arquivo)
    if padrao_br:
        dia, mes, ano = padrao_br.groups()
        try:
            return datetime(int(ano), int(mes), int(dia))
        except ValueError:
            pass

    padrao_iso = re.search(r'(\d{4})[-._](\d{2})[-._](\d{2})', nome_arquivo)
    if padrao_iso:
        ano, mes, dia = padrao_iso.groups()
        try:
            return datetime(int(ano), int(mes), int(dia))
        except ValueError:
            pass

    return None


def calcular_status_vencimento(dt_vencimento):
    """Mapeia a data para o status e nível de alerta visual."""
    if not dt_vencimento:
        return {"data_formatada": "Não identificada", "status": "Sem Data", "cor": "cinza"}
    
    hoje = datetime.now()
    dias_restantes = (dt_vencimento - hoje).days

    if dias_restantes < 0:
        return {"data_formatada": dt_vencimento.strftime("%d/%m/%Y"), "status": "Vencido", "cor": "vermelho"}
    elif dias_restantes <= 30:
        return {"data_formatada": dt_vencimento.strftime("%d/%m/%Y"), "status": "Atenção (Vence em breve)", "cor": "amarelo"}
    else:
        return {"data_formatada": dt_vencimento.strftime("%d/%m/%Y"), "status": "Em Dia", "cor": "verde"}


def identificar_tipo_doc_motorista(nome_arquivo):
    """Identifica a chave/categoria do documento do motorista para agrupamento dinâmico."""
    nome_norm = remover_acentos(nome_arquivo.lower())

    for chave, variacoes in MAPEAMENTO_MOTORISTAS.items():
        for v in variacoes:
            if remover_acentos(v.lower()) in nome_norm:
                return chave.upper()

    # Caso não esteja no mapeamento padrão, extrai o prefixo (Ex: "CDS", "OSD", "PAE", "TRM")
    partes = re.split(r'[-_\s]', nome_norm)
    if partes and len(partes[0]) >= 2:
        return partes[0].upper()

    return nome_arquivo.upper()


def identificar_tipo_doc_veiculo(nome_arquivo, caminho_pasta=""):
    """Identifica o tipo/categoria do documento do veículo para agrupamento."""
    nome_norm = remover_acentos(nome_arquivo.lower())
    caminho_norm = remover_acentos(caminho_pasta.lower())

    for chave in MAPEAMENTO_PASTAS_VEICULOS.keys():
        if chave in nome_norm or chave in caminho_norm:
            return chave.upper()
    
    siglas = ['CIV', 'CIPP', 'CRLV', 'CRONOTACOGRAFO', 'CRO', 'CALIBRAGEM', 'NF']
    for sigla in siglas:
        if sigla.lower() in nome_norm:
            return sigla
            
    return "OUTROS"


# --- CARREGAMENTO DA PLANILHA COLABORADORES.XLSX ---
def buscar_nome_por_cpf_na_planilha(cpf_digitado):
    if not os.path.exists(ARQUIVO_PLANILHA):
        print(f"[ERRO CRÍTICO] Planilha não encontrada em: {ARQUIVO_PLANILHA}")
        return None

    cpf_alvo_limpo = limpar_string(cpf_digitado)
    if not cpf_alvo_limpo:
        return None

    cpf_alvo_fmt = cpf_alvo_limpo.zfill(11)

    try:
        df = pd.read_excel(ARQUIVO_PLANILHA, dtype=str)

        col_cpf = None
        col_nome = None

        for col in df.columns:
            col_norm = remover_acentos(col).strip().upper()
            if 'CPF' in col_norm:
                col_cpf = col
            elif 'NOME' in col_norm or 'PASTA' in col_norm or 'COLABORADOR' in col_norm:
                col_nome = col

        if not col_cpf or not col_nome:
            print(f"[ERRO COLUNAS PLANILHA] Colunas detectadas: {list(df.columns)}")
            return None

        for _, row in df.iterrows():
            cpf_celula = limpar_string(row.get(col_cpf, ''))
            if not cpf_celula:
                continue

            if cpf_celula == cpf_alvo_limpo or cpf_celula.zfill(11) == cpf_alvo_fmt:
                nome_encontrado = str(row.get(col_nome, '')).strip()
                if nome_encontrado and nome_encontrado.lower() != 'nan':
                    print(f"\n[VALIDADO VIA CPF PLANILHA] CPF: {cpf_alvo_fmt} -> Nome/Pasta: '{nome_encontrado}'")
                    return nome_encontrado

    except Exception as e:
        print(f"[ERRO LEITURA PLANILHA] {e}")

    return None


# --- MÓDULO MOTORISTAS ---
def buscar_pasta_motorista(cpf_busca, nome_busca=""):
    print(f"\n==================================================")
    print(f"[BUSCA CONDUTOR] Digitado no site -> Nome: '{nome_busca}' | CPF: '{cpf_busca}'")

    nome_pasta_excel = buscar_nome_por_cpf_na_planilha(cpf_busca)

    if not nome_pasta_excel:
        print(f"[ERRO] O CPF '{cpf_busca}' não foi localizado na planilha Colaboradores.xlsx.")
        print(f"==================================================\n")
        return None

    nome_alvo_norm = remover_acentos(nome_pasta_excel).strip().lower()

    try:
        for raiz, dirs, _ in os.walk(PASTA_MOTORISTAS):
            for d in dirs:
                d_norm = remover_acentos(d).strip().lower()
                if d_norm == nome_alvo_norm or nome_alvo_norm in d_norm or d_norm in nome_alvo_norm:
                    caminho_completo = os.path.join(raiz, d)
                    print(f"[PASTA DA REDE ENCONTRADA] -> {caminho_completo}")
                    print(f"==================================================\n")
                    return caminho_completo

        caminho_direto = os.path.join(PASTA_MOTORISTAS, nome_pasta_excel)
        if os.path.exists(caminho_direto):
            return caminho_direto

    except Exception as e:
        print(f"[ERRO BUSCA NA REDE] {e}")

    print(f"[ERRO] A pasta '{nome_pasta_excel}' não existe no servidor {PASTA_MOTORISTAS}.")
    print(f"==================================================\n")
    return None


def buscar_arquivos_motorista(diretorio, documentos_solicitados):
    arquivos_encontrados = []
    if not os.path.exists(diretorio):
        return arquivos_encontrados

    for doc in documentos_solicitados:
        doc_limpo = remover_acentos(doc.lower().strip())
        palavras = re.findall(r'\b\w+\b', doc_limpo)

        termos = set()
        for palavra in palavras:
            for chave, variacoes in MAPEAMENTO_MOTORISTAS.items():
                if chave in palavra:
                    termos.update(variacoes)

        if not termos:
            termos.update([p for p in palavras if len(p) > 2])

        doc_encontrado = False
        for raiz, _, ficheiros in os.walk(diretorio):
            if doc_encontrado:
                break

            for ficheiro in ficheiros:
                nome_ficheiro_limpo = remover_acentos(ficheiro.lower())

                for termo in termos:
                    if remover_acentos(termo.lower()) in nome_ficheiro_limpo:
                        caminho_completo = os.path.join(raiz, ficheiro)
                        if caminho_completo not in arquivos_encontrados:
                            print(f"[DOCUMENTO LOCALIZADO] -> {ficheiro}")
                            arquivos_encontrados.append(caminho_completo)
                            doc_encontrado = True
                            break

    return arquivos_encontrados


# --- MÓDULO VEÍCULOS ---
def gerar_variacoes_placa(placa_bruta):
    placa_limpa = re.sub(r'[^A-Z0-9]', '', placa_bruta.upper().strip())
    if not placa_limpa:
        return []

    variacoes = {placa_limpa}

    if 'I' in placa_limpa:
        variacoes.add(placa_limpa.replace('I', '1'))
    if '1' in placa_limpa:
        variacoes.add(placa_limpa.replace('1', 'I'))

    if 'O' in placa_limpa:
        variacoes.add(placa_limpa.replace('O', '0'))
    if '0' in placa_limpa:
        variacoes.add(placa_limpa.replace('0', 'O'))

    return [v for v in variacoes if len(v) >= 5]


def extrair_ano_ou_data(nome_arquivo):
    anos = re.findall(r'20\d{2}', nome_arquivo)
    if anos:
        return max([int(a) for a in anos])
    return 0


def buscar_arquivos_veiculo(placa_busca, documentos_solicitados):
    arquivos_encontrados = []

    if not os.path.exists(PASTA_VEICULOS):
        print(f"[ERRO CRÍTICO] Caminho inacessível: {PASTA_VEICULOS}")
        return arquivos_encontrados

    variacoes_placa = gerar_variacoes_placa(placa_busca)

    for doc_item in documentos_solicitados:
        doc_norm = remover_acentos(str(doc_item).lower().strip())
        pastas_alvo = []

        for chave_doc, pastas_fisicas in MAPEAMENTO_PASTAS_VEICULOS.items():
            if chave_doc in doc_norm:
                pastas_alvo.extend(pastas_fisicas)

        if not pastas_alvo:
            pastas_alvo.append(str(doc_item).strip())

        candidatos_categoria = []

        for nome_pasta in pastas_alvo:
            caminho_pasta_doc = os.path.join(PASTA_VEICULOS, nome_pasta)

            if not os.path.exists(caminho_pasta_doc):
                continue

            for raiz, _, ficheiros in os.walk(caminho_pasta_doc):
                for ficheiro in ficheiros:
                    if not ficheiro.lower().endswith('.pdf'):
                        continue

                    nome_ficheiro_limpo = re.sub(r'[^A-Z0-9]', '', ficheiro.upper())

                    if any(v in nome_ficheiro_limpo for v in variacoes_placa if v):
                        caminho_completo = os.path.join(raiz, ficheiro)

                        try:
                            mtime = os.path.getmtime(caminho_completo)
                        except Exception:
                            mtime = 0

                        ano_nome = extrair_ano_ou_data(ficheiro)

                        candidatos_categoria.append({
                            'caminho': caminho_completo,
                            'nome': ficheiro,
                            'ano_nome': ano_nome,
                            'mtime': mtime
                        })

        if candidatos_categoria:
            candidatos_categoria.sort(key=lambda x: (x['ano_nome'], x['mtime']), reverse=True)
            mais_recente = candidatos_categoria[0]

            if mais_recente['caminho'] not in arquivos_encontrados:
                arquivos_encontrados.append(mais_recente['caminho'])

    return arquivos_encontrados


# --- MÓDULO LICENÇAS (REGRAS DE FILTRO REFINADAS) ---
def buscar_arquivos_licenca(uf_busca, documentos_solicitados):
    arquivos_encontrados = []

    if not os.path.exists(PASTA_LICENCAS):
        print(f"[ERRO CRÍTICO] Caminho inacessível: {PASTA_LICENCAS}")
        return arquivos_encontrados

    uf_normalizada = uf_busca.strip().upper()

    for doc_chave in documentos_solicitados:
        doc_chave_upper = doc_chave.upper().strip()

        nome_subpasta = MAPEAMENTO_PASTAS_LICENCAS.get(doc_chave, '')
        if not nome_subpasta:
            for k, v in MAPEAMENTO_PASTAS_LICENCAS.items():
                if k in doc_chave_upper or doc_chave_upper in k:
                    nome_subpasta = v
                    break

        caminho_pasta_alvo = os.path.join(PASTA_LICENCAS, nome_subpasta) if nome_subpasta else PASTA_LICENCAS

        if not os.path.exists(caminho_pasta_alvo):
            print(f"[AVISO LICENÇAS] Pasta não encontrada: {caminho_pasta_alvo}")
            continue

        candidatos_licenca = []

        for raiz, _, ficheiros in os.walk(caminho_pasta_alvo):
            for ficheiro in ficheiros:
                if not ficheiro.lower().endswith('.pdf'):
                    continue

                ficheiro_upper = remover_acentos(ficheiro.upper())

                if 'ANTT' in doc_chave_upper:
                    if 'ANTT' in ficheiro_upper:
                        caminho_completo = os.path.join(raiz, ficheiro)
                        mtime = os.path.getmtime(caminho_completo) if os.path.exists(caminho_completo) else 0
                        ano_nome = extrair_ano_ou_data(ficheiro)
                        candidatos_licenca.append({'caminho': caminho_completo, 'nome': ficheiro, 'ano_nome': ano_nome, 'mtime': mtime})
                    continue

                padrao_uf = rf'^\s*{uf_normalizada}\b|\b{uf_normalizada}\b'
                if re.search(padrao_uf, ficheiro_upper):

                    if 'CADASTRO' in doc_chave_upper or 'CTF' in doc_chave_upper:
                        if 'CADASTRO' not in ficheiro_upper and 'CTF' not in ficheiro_upper:
                            continue

                    elif 'IBAMA' in doc_chave_upper or 'AMBIENTAL' in doc_chave_upper:
                        if '1' in ficheiro_upper and 'PAG' in ficheiro_upper:
                            continue

                        if 'CADASTRO' in ficheiro_upper or 'CTF' in ficheiro_upper:
                            continue

                        if 'AUTORIZACAO AMBIENTAL' not in ficheiro_upper and 'AUTORIZAÇÃO AMBIENTAL' not in ficheiro_upper:
                            continue

                    caminho_completo = os.path.join(raiz, ficheiro)
                    try:
                        mtime = os.path.getmtime(caminho_completo)
                    except Exception:
                        mtime = 0

                    ano_nome = extrair_ano_ou_data(ficheiro)
                    candidatos_licenca.append({
                        'caminho': caminho_completo, 
                        'nome': ficheiro, 
                        'ano_nome': ano_nome, 
                        'mtime': mtime
                    })

        if candidatos_licenca:
            candidatos_licenca.sort(key=lambda x: (x['ano_nome'], x['mtime']), reverse=True)
            mais_recente = candidatos_licenca[0]

            if mais_recente['caminho'] not in arquivos_encontrados:
                print(f"[LICENÇA ENCONTRADA] Chave: {doc_chave} | UF: {uf_normalizada} -> {mais_recente['nome']}")
                arquivos_encontrados.append(mais_recente['caminho'])

    return arquivos_encontrados


# --- FUNÇÃO DE ENVIO VIA OUTLOOK LOCAL ---
def enviar_via_outlook(destinatario, assunto, corpo, anexos=[]):
    pythoncom.CoInitialize()
    try:
        try:
            outlook = win32.GetActiveObject('Outlook.Application')
        except Exception:
            outlook = win32.Dispatch('Outlook.Application')

        mail = outlook.CreateItem(0)
        mail.To = destinatario
        mail.Subject = assunto
        mail.Body = corpo

        for caminho_arquivo in anexos:
            if os.path.exists(caminho_arquivo):
                mail.Attachments.Add(caminho_arquivo)

        mail.Send()
        return True
    except Exception as e:
        raise Exception(f"Falha na integração com o Outlook local: {str(e)}")
    finally:
        pythoncom.CoUninitialize()


# --- ROTAS API ---

@app.before_request
def handle_preflight():
    if request.method == "OPTIONS":
        response = jsonify({'status': 'ok'})
        response.headers.add("Access-Control-Allow-Origin", "*")
        response.headers.add("Access-Control-Allow-Headers", "Content-Type,Authorization")
        response.headers.add("Access-Control-Allow-Methods", "GET,PUT,POST,DELETE,OPTIONS")
        return response, 200


@app.route('/send-email', methods=['POST'])
def send_email_direct():
    dados = request.json or {}
    to_email = dados.get('toEmail')
    subject = dados.get('subject', 'Notificação de Cadastro')
    body_content = dados.get('bodyContent', '<p>Segue a documentação referente ao cadastro.</p>')

    if not to_email:
        return jsonify({'success': False, 'message': 'O e-mail do destinatário é obrigatório.'}), 400

    try:
        msg = MIMEMultipart()
        msg['From'] = f"Henry Peressinotti <{EMAIL_USER}>"
        msg['To'] = to_email
        msg['Subject'] = subject
        msg.attach(MIMEText(body_content, 'html'))

        server = smtplib.SMTP('smtp.office365.com', 587)
        server.starttls()
        server.login(EMAIL_USER, EMAIL_PASS)
        server.send_message(msg)
        server.quit()

        return jsonify({'success': True, 'message': 'E-mail enviado com sucesso via SMTP!'}), 200

    except Exception as e:
        print(f"[ERRO SMTP] {e}")
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/enviar-motorista', methods=['POST'])
def enviar_docs_motorista():
    dados = request.json or {}
    nome_condutor = dados.get('nome', '').strip()
    cpf_digitado = dados.get('cpf', '').strip()
    docs_solicitados = dados.get('documentos', [])
    email_destinatario = dados.get('email', '').strip()

    if not docs_solicitados or not email_destinatario:
        return jsonify({'erro': 'Dados incompletos! Verifique o e-mail e os documentos.'}), 400

    pasta_condutor = buscar_pasta_motorista(cpf_digitado, nome_condutor)

    if not pasta_condutor:
        cpf_fmt = limpar_string(cpf_digitado)
        return jsonify({'erro': f'Atenção: Motorista (CPF {cpf_fmt}) não atualizou o documento solicitado ou nao esta na pasta.'}), 404

    anexos = buscar_arquivos_motorista(pasta_condutor, docs_solicitados)

    if not anexos:
        return jsonify({'erro': 'Nenhum dos documentos selecionados foi localizado na pasta deste motorista.'}), 404

    try:
        cpf_limpo_fmt = limpar_string(cpf_digitado)
        identificacao = f"{nome_condutor} (CPF: {cpf_limpo_fmt})" if nome_condutor else f"CPF: {cpf_limpo_fmt}"
        corpo_email = f"Olá,\n\nSegue em anexo a documentação solicitada do motorista: {identificacao}.\n\nDocumentos inclusos:\n"
        corpo_email += "\n".join([f"- {os.path.basename(a)}" for a in anexos])
        corpo_email += "\n\nAtenciosamente,\nPortal Transjordano"

        enviar_via_outlook(
            destinatario=email_destinatario,
            assunto=f"Documentação de Motorista - {identificacao}",
            corpo=corpo_email,
            anexos=anexos
        )
        return jsonify({'sucesso': True, 'mensagem': f'E-mail enviado automaticamente via Outlook com {len(anexos)} documento(s) anexado(s)!'})
    except Exception as e:
        return jsonify({'erro': str(e)}), 500


@app.route('/api/enviar-veiculo', methods=['POST'])
def enviar_docs_veiculo():
    dados = request.json or {}
    placa_digitada = dados.get('placa', '').strip().upper()
    docs_solicitados = dados.get('documentos', [])
    email_destinatario = dados.get('email', '').strip()

    if not placa_digitada or not docs_solicitados or not email_destinatario:
        return jsonify({'erro': 'Dados incompletos! Verifique a placa, e-mail e documentos selecionados.'}), 400

    anexos = buscar_arquivos_veiculo(placa_digitada, docs_solicitados)

    if not anexos:
        return jsonify({'erro': f'Nenhum documento localizado para a placa {placa_digitada} nas pastas selecionadas.'}), 404

    try:
        corpo_email = f"Olá,\n\nSegue em anexo a documentação solicitada do veículo placa {placa_digitada}.\n\nDocumentos inclusos:\n"
        corpo_email += "\n".join([f"- {os.path.basename(a)}" for a in anexos])
        corpo_email += "\n\nAtenciosamente,\nPortal Transjordano"

        enviar_via_outlook(
            destinatario=email_destinatario,
            assunto=f"Documentação Veicular ({placa_digitada}) - Transjordano",
            corpo=corpo_email,
            anexos=anexos
        )
        return jsonify({'sucesso': True, 'mensagem': f'E-mail enviado automaticamente via Outlook com {len(anexos)} documento(s) veicular(es) anexado(s)!'})
    except Exception as e:
        return jsonify({'erro': str(e)}), 500


@app.route('/api/enviar-licencas', methods=['POST'])
def enviar_docs_licencas():
    dados = request.json or {}
    estado_uf = dados.get('estado', '').strip().upper()
    docs_solicitados = dados.get('documentos', [])
    email_destinatario = dados.get('email', '').strip()

    if not estado_uf or not docs_solicitados or not email_destinatario:
        return jsonify({'erro': 'Dados incompletos! Verifique a UF, e-mail e licenças selecionadas.'}), 400

    anexos = buscar_arquivos_licenca(estado_uf, docs_solicitados)

    if not anexos:
        return jsonify({'erro': f'Nenhuma licença foi localizada para o estado {estado_uf} com os documentos selecionados.'}), 404

    corpo_email = f"Olá,\n\nSegue em anexo as licenças vigentes solicitadas do Estado: {estado_uf}.\n\nDocumentos inclusos:\n"
    corpo_email += "\n".join([f"- {os.path.basename(a)}" for a in anexos])
    corpo_email += "\n\nAtenciosamente,\nPortal Transjordano"

    try:
        enviar_via_outlook(
            destinatario=email_destinatario,
            assunto=f"Consulta de Licenças ({estado_uf}) - Transjordano",
            corpo=corpo_email,
            anexos=anexos
        )
        return jsonify({'sucesso': True, 'mensagem': f'E-mail enviado via Outlook com {len(anexos)} licença(s) atualizada(s)!'})
    except Exception as e_outlook:
        print(f"[AVISO OUTLOOK] Falha no Outlook local. Tentando envio direto via SMTP: {e_outlook}")

        try:
            msg = MIMEMultipart()
            msg['From'] = f"Henry Peressinotti <{EMAIL_USER}>"
            msg['To'] = email_destinatario
            msg['Subject'] = f"Consulta de Licenças ({estado_uf}) - Transjordano"
            msg.attach(MIMEText(corpo_email, 'plain'))

            for arq in anexos:
                if os.path.exists(arq):
                    with open(arq, 'rb') as f:
                        part = MIMEBase('application', 'octet-stream')
                        part.set_payload(f.read())
                        encoders.encode_base64(part)
                        part.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(arq)}"')
                        msg.attach(part)

            server = smtplib.SMTP('smtp.office365.com', 587)
            server.starttls()
            server.login(EMAIL_USER, EMAIL_PASS)
            server.send_message(msg)
            server.quit()

            return jsonify({'sucesso': True, 'mensagem': f'E-mail enviado via SMTP Office 365 com {len(anexos)} licença(s) atualizada(s)!'})
        except Exception as e_smtp:
            return jsonify({'erro': f'Erro ao enviar e-mail: {str(e_smtp)}'}), 500


# --- ROTAS DE CONSULTA DE VENCIMENTOS (AGRUPADAS PELO MAIS RECENTE) ---

@app.route('/api/vencimentos/motoristas', methods=['GET'])
def consultar_vencimentos_motoristas():
    cpf_busca = request.args.get('cpf', '').strip()
    nome_busca = request.args.get('nome', '').strip()

    if not cpf_busca:
        return jsonify({'erro': 'Informe o CPF para realizar a busca.'}), 400

    pasta_motorista = buscar_pasta_motorista(cpf_busca, nome_busca)
    if not pasta_motorista or not os.path.exists(pasta_motorista):
        return jsonify({'erro': f'Pasta não localizada para o CPF informado.'}), 404

    documentos_por_tipo = {}

    for raiz, _, ficheiros in os.walk(pasta_motorista):
        # Ignora arquivos de subpastas como HISTORICO
        if 'HISTORICO' in raiz.upper():
            continue

        for f in ficheiros:
            if f.lower().endswith(('.pdf', '.png', '.jpg', '.jpeg')):
                # Ignora arquivos auxiliares de CPF
                if f.upper().startswith('AA.CPF'):
                    continue

                caminho_completo = os.path.join(raiz, f)
                dt_venc = extrair_data_do_nome(f, caminho_completo)
                status_info = calcular_status_vencimento(dt_venc)
                tipo_doc = identificar_tipo_doc_motorista(f)

                try:
                    mtime = os.path.getmtime(caminho_completo)
                except Exception:
                    mtime = 0

                item = {
                    'documento': f,
                    'validade': status_info['data_formatada'],
                    'status': status_info['status'],
                    'cor': status_info['cor'],
                    'dt_venc': dt_venc,
                    'mtime': mtime
                }

                if tipo_doc not in documentos_por_tipo:
                    documentos_por_tipo[tipo_doc] = item
                else:
                    atual = documentos_por_tipo[tipo_doc]
                    dt_atual = atual['dt_venc']
                    dt_novo = item['dt_venc']

                    if dt_novo and dt_atual:
                        if dt_novo > dt_atual:
                            documentos_por_tipo[tipo_doc] = item
                    elif dt_novo and not dt_atual:
                        documentos_por_tipo[tipo_doc] = item
                    elif not dt_novo and not dt_atual:
                        if item['mtime'] > atual['mtime']:
                            documentos_por_tipo[tipo_doc] = item

    documentos_finais = list(documentos_por_tipo.values())

    return jsonify({
        'condutor': nome_busca or "Localizado via Planilha",
        'cpf': cpf_busca,
        'documentos': documentos_finais
    })


@app.route('/api/vencimentos/veiculos', methods=['GET'])
def consultar_vencimentos_veiculos():
    placa_busca = request.args.get('placa', '').strip().upper()

    if not placa_busca:
        return jsonify({'erro': 'Informe a placa para realizar a busca.'}), 400

    variacoes_placa = gerar_variacoes_placa(placa_busca)
    documentos_por_tipo = {}

    if os.path.exists(PASTA_VEICULOS):
        for raiz, _, ficheiros in os.walk(PASTA_VEICULOS):
            for f in ficheiros:
                if not f.lower().endswith('.pdf'):
                    continue
                
                nome_limpo = re.sub(r'[^A-Z0-9]', '', f.upper())
                if any(v in nome_limpo for v in variacoes_placa if v):
                    caminho_completo = os.path.join(raiz, f)
                    dt_venc = extrair_data_do_nome(f, caminho_completo)
                    status_info = calcular_status_vencimento(dt_venc)
                    tipo_doc = identificar_tipo_doc_veiculo(f, raiz)

                    try:
                        mtime = os.path.getmtime(caminho_completo)
                    except Exception:
                        mtime = 0

                    item = {
                        'documento': f,
                        'validade': status_info['data_formatada'],
                        'status': status_info['status'],
                        'cor': status_info['cor'],
                        'dt_venc': dt_venc,
                        'mtime': mtime
                    }

                    if tipo_doc not in documentos_por_tipo:
                        documentos_por_tipo[tipo_doc] = item
                    else:
                        atual = documentos_por_tipo[tipo_doc]
                        dt_atual = atual['dt_venc']
                        dt_novo = item['dt_venc']

                        if dt_novo and dt_atual:
                            if dt_novo > dt_atual:
                                documentos_por_tipo[tipo_doc] = item
                        elif dt_novo and not dt_atual:
                            documentos_por_tipo[tipo_doc] = item
                        elif not dt_novo and not dt_atual:
                            if item['mtime'] > atual['mtime']:
                                documentos_por_tipo[tipo_doc] = item

    documentos_finais = list(documentos_por_tipo.values())

    if not documentos_finais:
        return jsonify({'erro': f'Nenhum documento localizado para a placa {placa_busca}.'}), 404

    return jsonify({
        'placa': placa_busca,
        'documentos': documentos_finais
    })


if __name__ == '__main__':
    # Executa a aplicação Flask na porta 5000 acessível na rede
    app.run(host='0.0.0.0', port=5000, debug=True)