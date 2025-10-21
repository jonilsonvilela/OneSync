from flask import Flask, render_template, request, redirect, url_for, send_from_directory, flash, Response
from werkzeug.utils import secure_filename
from pathlib import Path
from datetime import datetime
import subprocess
import os
import threading
import json
import pandas as pd
import sys # <--- MODIFICAÇÃO: Necessário para sys.executable

app = Flask(__name__)
app.secret_key = 'secreta_onesync'

# --- ESTRUTURA DE DIRETÓRIOS ---
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / 'data'
LOGS_DIR = DATA_DIR / 'logs_execucao'
REPORTS_DIR = DATA_DIR / 'relatorios_sumarizados'
REPROCESS_DIR = DATA_DIR / 'reprocessamento' # <--- MODIFICAÇÃO: Adicionado

# Garante que as pastas existem
DATA_DIR.mkdir(parents=True, exist_ok=True)
LOGS_DIR.mkdir(parents=True, exist_ok=True)
REPORTS_DIR.mkdir(parents=True, exist_ok=True)
REPROCESS_DIR.mkdir(parents=True, exist_ok=True) # <--- MODIFICAÇÃO: Adicionado

# --- VARIÁVEIS GLOBAIS ---
process_output = []
process_running = False
process_status = {
    "total": 0,
    "concluidos": 0,
    "falhados": 0
}
current_batch_info = {"start": None, "end": None} # Guarda info do lote atual

# --- FUNÇÃO DE EXECUÇÃO DO RPA ---
# Modificada para aceitar start_row e end_row
def executar_script(start_row=None, end_row=None):
    """Executa o script RPA em uma thread separada, processando um lote específico."""
    global process_output, process_running, process_status, current_batch_info
    
    # --- DEBUG PRINT ADICIONADO ---
    print(f"DEBUG (app.py - executar_script): Thread iniciada com start_row='{start_row}' (tipo: {type(start_row)}), end_row='{end_row}' (tipo: {type(end_row)})")
    # --- FIM DEBUG PRINT ---

    process_output, process_running = [], True
    process_status.update({"concluidos": 0, "falhados": 0, "total": 0})
    # Atualiza current_batch_info aqui também para consistência
    current_batch_info = {"start": start_row, "end": end_row} 

    # Limpa os caminhos de arquivo (mantido)
    app.config['LAST_LOG_FILE'] = None
    app.config['LAST_REPORT_FILE'] = None
    app.config['LAST_REPROCESS_FILE'] = None

    try:
        # Lê a planilha para determinar o tamanho (mantido)
        try:
            df_full = pd.read_excel(DATA_DIR / 'entrada.xlsx', sheet_name="Dados")
            total_rows_in_file = len(df_full)
            start_row_int = int(start_row) if start_row and start_row.isdigit() else None # Converte apenas se for dígito
            end_row_int = int(end_row) if end_row and end_row.isdigit() else None   # Converte apenas se for dígito
            start_index = start_row_int - 2 if start_row_int else 0
            end_index = end_row_int - 2 if end_row_int else total_rows_in_file - 1
            start_index = max(0, start_index)
            end_index = min(total_rows_in_file - 1, end_index)
            if start_index > end_index: process_status["total"] = 0 
            else: process_status["total"] = (end_index - start_index) + 1
        except Exception as e:
            process_output.append(f"[ERRO] Falha ao ler planilha para determinar tamanho do lote: {e}")
            process_status["total"] = 0 
            
        process_output.append(f"[INFO] Iniciando lote: Linha Excel {start_row or 'Início'} até {end_row or 'Fim'} (Total: {process_status['total']} processos)")

        # Monta o comando
        comando = [sys.executable, '-u', str(BASE_DIR / 'rpa_atualizar_pasta.py')]
        
        # --- DEBUG PRINT ADICIONADO ---
        print(f"DEBUG (app.py - executar_script): Verificando start_row ('{start_row}')")
        # --- FIM DEBUG PRINT ---
        # Verifica se start_row existe, não é None E NÃO é uma string vazia/espaços
        if start_row is not None and str(start_row).strip(): 
            comando.extend(['--start', str(start_row)])
            print(f"DEBUG (app.py - executar_script): Adicionado --start {start_row}") # Debug
        else:
             print(f"DEBUG (app.py - executar_script): --start NÃO adicionado.") # Debug

        # --- DEBUG PRINT ADICIONADO ---
        print(f"DEBUG (app.py - executar_script): Verificando end_row ('{end_row}')")
        # --- FIM DEBUG PRINT ---
         # Verifica se end_row existe, não é None E NÃO é uma string vazia/espaços
        if end_row is not None and str(end_row).strip():
            comando.extend(['--end', str(end_row)])
            print(f"DEBUG (app.py - executar_script): Adicionado --end {end_row}") # Debug
        else:
            print(f"DEBUG (app.py - executar_script): --end NÃO adicionado.") # Debug

        # --- DEBUG PRINT ADICIONADO ---
        print(f"DEBUG (app.py - executar_script): Comando subprocess final: {' '.join(comando)}") 
        # --- FIM DEBUG PRINT ---

        p = subprocess.Popen(
            comando, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, 
            text=True, encoding='utf-8', errors='replace', cwd=BASE_DIR,
            env=os.environ.copy()
        )
        
        # Leitura da saída do processo (sem alterações na lógica de leitura e contagem)
        for linha in iter(p.stdout.readline, ''):
            texto = linha.strip()
            if texto:
                process_output.append(texto)
                if "FINAL_PATHS:" in texto:
                    try:
                        paths_json = texto.split("FINAL_PATHS:", 1)[1]
                        paths = json.loads(paths_json)
                        if paths.get("json_log"): app.config['LAST_LOG_FILE'] = Path(paths["json_log"]).name
                        if paths.get("csv_report"): app.config['LAST_REPORT_FILE'] = Path(paths["csv_report"]).name
                        if paths.get("xlsx_reprocess"): app.config['LAST_REPROCESS_FILE'] = Path(paths["xlsx_reprocess"]).name
                    except Exception as e:
                        print(f"Erro ao parsear linha FINAL_PATHS: {e}")
                        process_output.append(f"[ERRO INTERNO APP.PY] Falha ao parsear caminhos finais: {e}")

                if "[SALVO] Processo" in texto: process_status["concluidos"] += 1
                elif "[ERRO]" in texto or "[ERRO CRÍTICO]" in texto or "[ERRO INESPERADO]" in texto:
                     is_final_error_line = "[INFO] Processamento linha" in texto and "concluído" in texto
                     if is_final_error_line and not any("[SALVO]" in log_line for log_line in process_output[-5:]): process_status["falhados"] += 1
                     elif "[ERRO] Falha crítica salvar" in texto or "[ERRO GRAVE] Conexão/Navegador perdido" in texto: process_status["falhados"] += 1

        p.stdout.close()
        ret_code = p.wait() 
        if ret_code != 0: process_output.append(f"[AVISO] O script RPA terminou com código de erro: {ret_code}")
        process_output.append("[INFO] Fim da execução do lote.")
        
    except Exception as e:
        process_output.append(f"[ERRO CRÍTICO] Falha ao executar o script RPA: {e}")
    finally:
        process_running = False

# --- FUNÇÃO AUXILIAR PARA PROCESSAR DADOS DO DASHBOARD ---
# (Seu código original, sem alterações)
def processar_dados_dashboard(dados_log):
    total_processos = len(dados_log)
    
    sucessos = sum(1 for p in dados_log if p.get("Status Geral") == "Sucesso")
    falhas = sum(1 for p in dados_log if p.get("Status Geral") == "Falha")
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
    # <--- MODIFICAÇÃO: Usa 'file' para corresponder ao HTML 'index.html' ---
    arquivo = request.files.get('file') 
    if not arquivo or not arquivo.filename.endswith('.xlsx'):
        flash('Nenhum arquivo .xlsx válido foi enviado.')
        return redirect(url_for('index'))

    upload_path = DATA_DIR / 'entrada_original.xlsx'
    arquivo.save(upload_path)
    
    try:
        import shutil
        shutil.copyfile(upload_path, DATA_DIR / 'entrada.xlsx')
        flash(f'Arquivo "{secure_filename(arquivo.filename)}" enviado e copiado para "entrada.xlsx".')
    except Exception as e:
         flash(f'Erro ao copiar arquivo para entrada.xlsx: {e}')
         
    return redirect(url_for('index'))

@app.route('/executar') 
def executar_rpa():
    if not (DATA_DIR / 'entrada.xlsx').exists():
        flash('Faça o upload do arquivo "entrada.xlsx" primeiro.')
        return redirect(url_for('index'))

    # --- CORREÇÃO: Usar 'start' e 'end' ---
    start_row = request.args.get('start') 
    end_row = request.args.get('end')
    # --- FIM DA CORREÇÃO ---

    # --- DEBUG PRINT (Mantido para confirmação) ---
    print(f"DEBUG (app.py - rota /executar): Recebido start='{start_row}' (tipo: {type(start_row)}), end='{end_row}' (tipo: {type(end_row)})") 
    # --- FIM DEBUG PRINT ---

    # Validação (Ajustada para os nomes corretos)
    try:
        if start_row and not start_row.isdigit(): raise ValueError("start não é dígito")
        if end_row and not end_row.isdigit(): raise ValueError("end não é dígito")
    except ValueError as e:
        print(f"DEBUG (app.py - rota /executar): Erro de validação - {e}") 
        flash(f'Linha inicial ou final inválida ({e}). Use apenas números.')
        return redirect(url_for('index'))

    if not process_running:
         print(f"DEBUG (app.py - rota /executar): Iniciando thread com start_row='{start_row}', end_row='{end_row}'") 
         # Inicia a thread passando os parâmetros corretos
         threading.Thread(target=executar_script, args=(start_row, end_row)).start()
    else:
         # Informa sobre o lote em execução (mantido)
         info = current_batch_info
         msg = "Processo já em execução."
         if info.get('start') or info.get('end'):
              msg += f" (Lote atual: Linha {info.get('start') or 'Início'} até {info.get('end') or 'Fim'})"
         flash(msg)

    # Passa a info do lote para o template 'executar.html' (mantido)
    # A variável current_batch_info ainda usa 'start'/'end' internamente, isso está OK.
    return render_template('executar.html', batch_info={"start": start_row, "end": end_row})

@app.route('/stream')
def stream():
    def event_stream():
        import time
        last_index = 0
        while process_running or last_index < len(process_output):
            current_len = len(process_output)
            if last_index < current_len:
                for i in range(last_index, current_len):
                    yield f"data: {json.dumps({'type': 'log', 'message': process_output[i]})}\n\n"
                last_index = current_len

            restantes = process_status["total"] - (process_status["concluidos"] + process_status["falhados"])
            progress_data = {'type': 'progress', **process_status, 'restantes': max(0, restantes)}
            yield f"data: {json.dumps(progress_data)}\n\n"

            if not process_running and last_index == len(process_output):
                # <--- MODIFICAÇÃO: Envia os caminhos dos arquivos no final do stream ---
                paths = {}
                # Usa as rotas _recente para os links
                if app.config.get('LAST_LOG_FILE'):
                    paths['json_log_url'] = url_for('baixar_log_recente') 
                if app.config.get('LAST_REPORT_FILE'):
                    paths['csv_report_url'] = url_for('baixar_relatorio_recente')
                if app.config.get('LAST_REPROCESS_FILE'):
                    paths['xlsx_reprocess_url'] = url_for('baixar_reprocess_recente')
                
                if paths:
                    yield f"data: {json.dumps({'type': 'final_paths', 'paths': paths})}\n\n"
                # --- FIM DA MODIFICAÇÃO ---
                break 

            time.sleep(0.5) 
    return Response(event_stream(), mimetype="text/event-stream")


@app.route('/resultado')
def resultado():
    # (Seu código original, sem alterações)
    logs = sorted(LOGS_DIR.glob('log_*.json'), key=os.path.getmtime, reverse=True)
    if not logs:
        return render_template('resultado.html', dashboard_data=None)

    try:
        with open(logs[0], 'r', encoding='utf-8') as f:
            dados_log = json.load(f)
        
        dashboard_data = processar_dados_dashboard(dados_log)
        dashboard_data['nome_arquivo'] = logs[0].name
        return render_template('resultado.html', dashboard_data=dashboard_data)
    except json.JSONDecodeError:
         flash(f"Erro ao ler o último arquivo de log ({logs[0].name}): formato JSON inválido.")
         return render_template('resultado.html', dashboard_data=None)
    except Exception as e:
         flash(f"Erro inesperado ao carregar resultado: {e}")
         return render_template('resultado.html', dashboard_data=None)


# --- ROTAS DE HISTÓRICO E RELATÓRIOS ---
# (Seu código original, com adição da rota de reprocessamento)
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

# <--- MODIFICAÇÃO: Nova rota para histórico de reprocessamento --->
@app.route('/historico_reprocessamento')
def historico_reprocessamento():
    arquivos = listar_arquivos(REPROCESS_DIR, 'xlsx', 'falhas_', 'Falhas de %d/%m/%Y às %H:%M:%S')
    # Assume que você criou o template 'historico_reprocessamento.html'
    return render_template('historico_reprocessamento.html', arquivos=arquivos)
# --- FIM DA MODIFICAÇÃO ---

@app.route('/ver_log/<path:nome_arquivo>')
def ver_log(nome_arquivo):
    # (Seu código original, sem alterações)
    path = LOGS_DIR / secure_filename(nome_arquivo) 
    if not path.is_file():
         flash(f"Arquivo de log '{nome_arquivo}' não encontrado.")
         return redirect(url_for('historico_logs'))
    try:
        with open(path, 'r', encoding='utf-8') as f:
            dados_log = json.load(f)
        dashboard_data = processar_dados_dashboard(dados_log)
        dashboard_data['nome_arquivo'] = nome_arquivo
        return render_template('resultado.html', dashboard_data=dashboard_data)
    except json.JSONDecodeError:
        flash(f"Erro ao ler o log '{nome_arquivo}': formato JSON inválido.")
        return redirect(url_for('historico_logs'))
    except Exception as e:
        flash(f"Erro ao carregar log '{nome_arquivo}': {e}")
        return redirect(url_for('historico_logs'))

@app.route('/ver_relatorio_csv/<path:nome_arquivo>')
def ver_relatorio_csv(nome_arquivo):
    # (Seu código original, sem alterações)
    path = REPORTS_DIR / secure_filename(nome_arquivo) 
    if not path.is_file():
         flash(f"Arquivo de relatório '{nome_arquivo}' não encontrado.")
         return redirect(url_for('historico_relatorios'))
    try:
        df = pd.read_csv(path, sep=';', encoding='utf-8-sig')
        max_rows_display = 500 
        table_html = df.head(max_rows_display).to_html(classes='table-auto w-full text-left', index=False)
        if len(df) > max_rows_display:
             table_html += f"<p class='mt-4 text-sm text-yellow-600'>Aviso: Exibindo apenas as primeiras {max_rows_display} linhas de {len(df)}. Baixe o CSV para ver o relatório completo.</p>"
        return render_template('ver_relatorio_csv.html', nome_arquivo=nome_arquivo, tabela=table_html)
    except Exception as e:
        flash(f"Erro ao ler o relatório '{nome_arquivo}': {e}")
        return redirect(url_for('historico_relatorios'))

# --- ROTAS DE DOWNLOAD ---
@app.route('/baixar_log/<path:nome>')
def baixar_log(nome):
    # (Seu código original, sem alterações)
    return send_from_directory(LOGS_DIR, secure_filename(nome), as_attachment=True)

@app.route('/baixar_relatorio_csv/<path:nome>')
def baixar_relatorio_csv(nome):
    # (Seu código original, sem alterações)
    return send_from_directory(REPORTS_DIR, secure_filename(nome), as_attachment=True)

# <--- MODIFICAÇÃO: Nova rota para baixar arquivo de reprocessamento do histórico --->
@app.route('/baixar_reprocessamento_xlsx/<path:nome>')
def baixar_reprocessamento_xlsx(nome):
    return send_from_directory(REPROCESS_DIR, secure_filename(nome), as_attachment=True)
# --- FIM DA MODIFICAÇÃO ---

# <--- MODIFICAÇÃO: Novas rotas para baixar arquivos "recentes" da página de execução --->
@app.route('/baixar_log_recente')
def baixar_log_recente():
    nome_arquivo = app.config.get('LAST_LOG_FILE') # Usa .get() para evitar erro se não existir
    if not nome_arquivo:
        flash("Nenhum log recente encontrado.")
        # Redireciona de volta para a página de execução com os parâmetros do lote
        return redirect(url_for('executar_rpa', **current_batch_info)) 
    return send_from_directory(LOGS_DIR, secure_filename(nome_arquivo), as_attachment=True)

@app.route('/baixar_relatorio_recente')
def baixar_relatorio_recente():
    nome_arquivo = app.config.get('LAST_REPORT_FILE')
    if not nome_arquivo:
        flash("Nenhum relatório recente encontrado.")
        return redirect(url_for('executar_rpa', **current_batch_info))
    return send_from_directory(REPORTS_DIR, secure_filename(nome_arquivo), as_attachment=True)

@app.route('/baixar_reprocess_recente')
def baixar_reprocess_recente():
    nome_arquivo = app.config.get('LAST_REPROCESS_FILE')
    if not nome_arquivo:
        flash("Nenhum arquivo de reprocessamento encontrado (provavelmente não houve falhas).")
        return redirect(url_for('executar_rpa', **current_batch_info))
    return send_from_directory(REPROCESS_DIR, secure_filename(nome_arquivo), as_attachment=True)
# --- FIM DA MODIFICAÇÃO ---


if __name__ == '__main__':
    # (Seu código original, sem alterações)
    app.run(debug=True, host='0.0.0.0', port=8080)