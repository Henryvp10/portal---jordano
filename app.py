import os
import re
import unicodedata
import pythoncom
import win32com.client as win32
import smtplib
import pandas as pd
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from flask import Flask, request, jsonify
from flask_cors import CORS

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
    'toxicologico': ['toxicologico', 'toxico'],
    'psicologica': ['psicologica', 'av.psicologica', 'avaliacao', 'avaliacaopsicologica', 'psico']
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

                # 1. Trata licença ANTT
                if 'ANTT' in doc_chave_upper:
                    if 'ANTT' in ficheiro_upper:
                        caminho_completo = os.path.join(raiz, ficheiro)
                        mtime = os.path.getmtime(caminho_completo) if os.path.exists(caminho_completo) else 0
                        ano_nome = extrair_ano_ou_data(ficheiro)
                        candidatos_licenca.append({'caminho': caminho_completo, 'nome': ficheiro, 'ano_nome': ano_nome, 'mtime': mtime})
                    continue

                # 2. Verifica se o arquivo pertence ao Estado (UF) selecionado
                padrao_uf = rf'^\s*{uf_normalizada}\b|\b{uf_normalizada}\b'
                if re.search(padrao_uf, ficheiro_upper):

                    # --- FILTROS RIGOROSOS DE LICENÇAS ---

                    # Regra para CADASTRO FEDERAL / CTF
                    if 'CADASTRO' in doc_chave_upper or 'CTF' in doc_chave_upper:
                        if 'CADASTRO' not in ficheiro_upper and 'CTF' not in ficheiro_upper:
                            continue

                    # Regra para IBAMA / AMBIENTAL COMPLETO
                    elif 'IBAMA' in doc_chave_upper or 'AMBIENTAL' in doc_chave_upper:
                        # Ignora resumos/páginas únicas como "1ª PAG" ou "1 PAG"
                        if '1' in ficheiro_upper and 'PAG' in ficheiro_upper:
                            continue

                        # Garante que não é Cadastro Técnico Federal
                        if 'CADASTRO' in ficheiro_upper or 'CTF' in ficheiro_upper:
                            continue

                        # Exige que seja especificamente o documento completo "AUTORIZAÇÃO AMBIENTAL"
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

        # ORDENAÇÃO POR DATA: Retorna o documento mais recente encontrado
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


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)