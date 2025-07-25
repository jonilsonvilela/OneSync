from flask import Flask, render_template, request, redirect, url_for, send_from_directory, flash, Response
from werkzeug.utils import secure_filename
from pathlib import Path
from datetime import datetime
import subprocess
import os
import threading
import json
import pandas as pd

app = Flask(__name__)
app.secret_key = 'secreta_onesync'

# --- ESTRUTURA DE DIRETÓRIOS ---
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / 'data'
LOGS_DIR = DATA_DIR / 'logs_execucao'
REPORTS_DIR = DATA_DIR / 'relatorios_sumarizados' # <-- Nova pasta de relatórios

# Garante que as pastas existem
DATA_DIR.mkdir(parents=True, exist_ok=True)
LOGS_DIR.mkdir(parents=True, exist_ok=True)
REPORTS_DIR.mkdir(parents=True, exist_ok=True)

# --- VARIÁVEIS GLOBAIS ---
process_output = []
process_running = False
process_status = {
    "total": 0,
    "concluidos": 0,
    "falhados": 0
}

# --- FUNÇÃO DE EXECUÇÃO DO RPA ---
def executar_script():
    """Executa o script RPA em uma thread separada."""
    global process_output, process_running, process_status
    process_output, process_running = [], True
    process_status.update({"concluidos": 0, "falhados": 0})
    
    try:
        comando = ['python', '-u', str(BASE_DIR / 'rpa_atualizar_pasta.py')]
        p = subprocess.Popen(
            comando, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            text=True, encoding='utf-8', errors='replace', cwd=BASE_DIR
        )
        for linha in iter(p.stdout.readline, ''):
            texto = linha.strip()
            if texto:
                process_output.append(texto)
                if "[SALVO] Processo" in texto: process_status["concluidos"] += 1
                elif "[Acesso ao Processo] Falha" in texto or "[ERRO] Falha ao salvar" in texto:
                    process_status["falhados"] += 1
        p.stdout.close()
        p.wait()
    except Exception as e:
        process_output.append(f"[ERRO CRÍTICO] Falha ao executar o script RPA: {e}")
    finally:
        process_running = False

# --- FUNÇÃO AUXILIAR PARA PROCESSAR DADOS DO DASHBOARD ---
def processar_dados_dashboard(dados_log):
    total_processos = len(dados_log)
    sucessos = sum(1 for p in dados_log if p.get("Status Geral") == "Sucesso")
    falhas = total_processos - sucessos
    tempo_total = sum(d for p in dados_log for d in p.get("Duracao por etapa (s)", {}).values())
    
    duracao_por_etapa = {}
    for processo in dados_log:
        for etapa, duracao in processo.get("Duracao por etapa (s)", {}).items():
            duracao_por_etapa[etapa] = duracao_por_etapa.get(etapa, 0) + duracao
    
    etapas_ordenadas = sorted(duracao_por_etapa.items(), key=lambda item: item[1], reverse=True)
    
    return {
        'total_processos': total_processos,
        'sucessos': sucessos,
        'falhas': falhas,
        'percentual_sucesso': (sucessos / total_processos * 100) if total_processos > 0 else 0,
        'percentual_falha': (falhas / total_processos * 100) if total_processos > 0 else 0,
        'tempo_total_execucao': tempo_total,
        'grafico_duracao': {
            'labels': [e for e, d in etapas_ordenadas],
            'data': [round(d, 2) for e, d in etapas_ordenadas]
        },
        'dados_brutos': dados_log
    }

# --- ROTAS PRINCIPAIS ---
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    arquivo = request.files.get('arquivo')
    if not arquivo or not arquivo.filename.endswith('.xlsx'):
        flash('Nenhum arquivo .xlsx válido foi enviado.')
        return redirect(url_for('index'))

    arquivo.save(DATA_DIR / 'entrada.xlsx')
    flash(f'Arquivo "{secure_filename(arquivo.filename)}" enviado com sucesso.')
    return redirect(url_for('index'))

@app.route('/executar')
def executar_rpa():
    if not (DATA_DIR / 'entrada.xlsx').exists():
        flash('Faça o upload do arquivo "entrada.xlsx" primeiro.')
        return redirect(url_for('index'))

    if not process_running:
        try:
            df = pd.read_excel(DATA_DIR / 'entrada.xlsx', sheet_name="Dados")
            process_status["total"] = len(df)
            threading.Thread(target=executar_script).start()
        except Exception as e:
            flash(f"Erro ao ler a planilha: {e}")
            return redirect(url_for('index'))
            
    return render_template('executar.html')

@app.route('/stream')
def stream():
    def event_stream():
        import time
        last_index = 0
        while process_running or last_index < len(process_output):
            while last_index < len(process_output):
                yield f"data: {json.dumps({'type': 'log', 'message': process_output[last_index]})}\n\n"
                last_index += 1
            
            restantes = process_status["total"] - (process_status["concluidos"] + process_status["falhados"])
            progress_data = {'type': 'progress', **process_status, 'restantes': max(0, restantes)}
            yield f"data: {json.dumps(progress_data)}\n\n"
            
            if not process_running: break
            time.sleep(0.5)
    return Response(event_stream(), mimetype="text/event-stream")

@app.route('/resultado')
def resultado():
    logs = sorted(LOGS_DIR.glob('log_*.json'), key=os.path.getmtime, reverse=True)
    if not logs:
        return render_template('resultado.html', dashboard_data=None)

    with open(logs[0], 'r', encoding='utf-8') as f:
        dados_log = json.load(f)
    
    dashboard_data = processar_dados_dashboard(dados_log)
    dashboard_data['nome_arquivo'] = logs[0].name
    return render_template('resultado.html', dashboard_data=dashboard_data)

# --- ROTAS DE HISTÓRICO E RELATÓRIOS ---
def listar_arquivos(diretorio, extensao, prefixo, formato_data):
    arquivos = sorted(diretorio.glob(f'{prefixo}*.{extensao}'), key=os.path.getmtime, reverse=True)
    lista_formatada = []
    for arq in arquivos:
        try:
            ts_str = arq.name.replace(prefixo, '').replace(f'.{extensao}', '')
            dt_obj = datetime.strptime(ts_str, '%Y-%m-%d_%H-%M-%S')
            nome_exibicao = dt_obj.strftime(formato_data)
        except ValueError:
            nome_exibicao = arq.name
        lista_formatada.append({'original': arq.name, 'exibicao': nome_exibicao})
    return lista_formatada

@app.route('/historico_logs')
def historico_logs():
    arquivos = listar_arquivos(LOGS_DIR, 'json', 'log_', 'Log de %d/%m/%Y às %H:%M:%S')
    return render_template('historico_logs.html', arquivos=arquivos)

@app.route('/historico_relatorios')
def historico_relatorios():
    arquivos = listar_arquivos(REPORTS_DIR, 'csv', 'relatorio_', 'Relatório de %d/%m/%Y às %H:%M:%S')
    return render_template('historico_relatorios.html', arquivos=arquivos)

@app.route('/ver_log/<path:nome_arquivo>')
def ver_log(nome_arquivo):
    try:
        with open(LOGS_DIR / nome_arquivo, 'r', encoding='utf-8') as f:
            dados_log = json.load(f)
        dashboard_data = processar_dados_dashboard(dados_log)
        dashboard_data['nome_arquivo'] = nome_arquivo
        return render_template('resultado.html', dashboard_data=dashboard_data)
    except Exception as e:
        flash(f"Erro ao ler o log '{nome_arquivo}': {e}")
        return redirect(url_for('historico_logs'))

@app.route('/ver_relatorio_csv/<path:nome_arquivo>')
def ver_relatorio_csv(nome_arquivo):
    try:
        df = pd.read_csv(REPORTS_DIR / nome_arquivo, sep=';', encoding='utf-8-sig')
        return render_template('ver_relatorio_csv.html', nome_arquivo=nome_arquivo, tabela=df.to_html(classes='table-auto w-full text-left', index=False))
    except Exception as e:
        flash(f"Erro ao ler o relatório '{nome_arquivo}': {e}")
        return redirect(url_for('historico_relatorios'))

@app.route('/baixar_log/<path:nome>')
def baixar_log(nome):
    return send_from_directory(LOGS_DIR, nome, as_attachment=True)

@app.route('/baixar_relatorio_csv/<path:nome>')
def baixar_relatorio_csv(nome):
    return send_from_directory(REPORTS_DIR, nome, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True, port=8080)