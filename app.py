from flask import Flask, render_template, request, redirect, url_for, send_from_directory, flash, Response
from werkzeug.utils import secure_filename
from pathlib import Path
import subprocess
import os
import threading

app = Flask(__name__)
app.secret_key = 'secreta_onesync'

# Diretórios base
BASE_DIR = Path(__file__).resolve().parent
UPLOAD_FOLDER = BASE_DIR / 'uploads'
DATA_DIR = BASE_DIR / 'data'
RELATORIOS_DIR = DATA_DIR / 'relatorios'

# Garante que as pastas existem
UPLOAD_FOLDER.mkdir(parents=True, exist_ok=True)
RELATORIOS_DIR.mkdir(parents=True, exist_ok=True)

# Variáveis globais para log ao vivo
process_output = []
process_running = False

def executar_script():
    global process_output, process_running
    process_output = []
    process_running = True
    try:
        print("[LOG] Iniciando subprocesso...", flush=True)

        p = subprocess.Popen(
            ['python', '-u', 'rpa_atualizar_pasta.py'],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,  # Equivalente moderno do universal_newlines
            encoding='utf-8',  # Especifica a codificação esperada
            errors='replace' # Evita que a aplicação quebre se algum caractere inválido aparecer
        )

        while True:
            linha = p.stdout.readline()
            if linha == '' and p.poll() is not None:
                break
            if linha:
                texto = linha.rstrip()
                print(f"[LOG] {texto}", flush=True)
                process_output.append(texto)

        p.wait()

    except Exception as e:
        erro = f"[ERRO] {str(e)}"
        print(erro, flush=True)
        process_output.append(erro)

    process_running = False

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    arquivo = request.files.get('arquivo')
    if not arquivo or arquivo.filename == '':
        flash('Nenhum arquivo enviado.')
        return redirect(url_for('index'))

    if not arquivo.filename.endswith('.xlsx'):
        flash('Apenas arquivos .xlsx são aceitos.')
        return redirect(url_for('index'))

    # Salva tanto na pasta de uploads quanto na pasta data (para leitura do RPA)
    upload_path = UPLOAD_FOLDER / 'entrada.xlsx'
    data_path = DATA_DIR / 'entrada.xlsx'
    arquivo.save(upload_path)
    arquivo.save(data_path)

    flash('Arquivo enviado com sucesso.')
    return redirect(url_for('index'))

@app.route('/executar_stream')
def executar_stream():
    global process_running
    if not process_running:
        thread = threading.Thread(target=executar_script)
        thread.start()
    return render_template('executar.html')

@app.route('/stream')
def stream():
    def event_stream():
        import time
        last_index = 0
        while process_running or last_index < len(process_output):
            while last_index < len(process_output):
                yield f"data: {process_output[last_index]}\n\n"
                last_index += 1
            time.sleep(0.3)  # 🔁 força checagem periódica para liberar o buffer
    return Response(event_stream(), mimetype="text/event-stream")

@app.route('/executar', methods=['POST'])
def executar():
    try:
        resultado = subprocess.run(
            ['python', 'rpa_atualizar_pasta.py'],
            cwd=BASE_DIR,
            capture_output=True,
            text=True,
            timeout=600
        )
        with open(RELATORIOS_DIR / 'ultimo_output.txt', 'w', encoding='utf-8') as f:
            f.write(resultado.stdout + "\n" + resultado.stderr)
        flash('Script executado. Verifique o relatório.')
    except Exception as e:
        flash(f'Erro na execução: {e}')
    return redirect(url_for('resultado'))

@app.route('/resultado')
def resultado():
    relatorios = sorted(RELATORIOS_DIR.glob('relatorio_atualizacao*.txt'), reverse=True)
    ultimo = relatorios[0] if relatorios else None
    conteudo = ultimo.read_text(encoding='utf-8') if ultimo else "Nenhum relatório encontrado."
    return render_template('resultado.html', relatorio=conteudo)

@app.route('/historico')
def historico():
    arquivos = sorted(RELATORIOS_DIR.glob('relatorio_atualizacao*.txt'), reverse=True)
    return render_template('historico.html', arquivos=arquivos)

@app.route('/relatorios/<path:nome>')
def baixar_relatorio(nome):
    return send_from_directory(RELATORIOS_DIR, nome, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True, port=8080)
