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
current_batch_info = {"start": None, "end": None} # Guarda info do lote atual

# --- FUNÇÃO DE EXECUÇÃO DO RPA ---
# Modificada para aceitar start_row e end_row
def executar_script(start_row=None, end_row=None):
    """Executa o script RPA em uma thread separada, processando um lote específico."""
    global process_output, process_running, process_status, current_batch_info
    process_output, process_running = [], True
    # Resetar contadores para o novo lote
    process_status.update({"concluidos": 0, "falhados": 0, "total": 0})
    current_batch_info = {"start": start_row, "end": end_row}

    try:
        # Lê a planilha para determinar o tamanho real do lote
        try:
            df_full = pd.read_excel(DATA_DIR / 'entrada.xlsx', sheet_name="Dados")
            total_rows_in_file = len(df_full)
            
            # Converte linhas Excel (base 1) para índices pandas (base 0)
            # Linha 2 Excel -> índice 0 pandas
            start_index = int(start_row) - 2 if start_row else 0
            end_index = int(end_row) - 2 if end_row else total_rows_in_file - 1
            
            # Garante que os índices sejam válidos
            start_index = max(0, start_index)
            end_index = min(total_rows_in_file - 1, end_index)

            if start_index > end_index:
                 process_status["total"] = 0 # Lote inválido ou vazio
            else:
                 process_status["total"] = (end_index - start_index) + 1
                 
        except Exception as e:
            process_output.append(f"[ERRO] Falha ao ler planilha para determinar tamanho do lote: {e}")
            process_status["total"] = 0 # Não é possível determinar o total
            # Considerar parar a execução aqui se a leitura falhar?
            
        process_output.append(f"[INFO] Iniciando lote: Linha Excel {start_row or 'Início'} até {end_row or 'Fim'} (Total: {process_status['total']} processos)")

        # Monta o comando com os argumentos de lote
        comando = ['python', '-u', str(BASE_DIR / 'rpa_atualizar_pasta.py')]
        if start_row:
            comando.extend(['--start', str(start_row)])
        if end_row:
            comando.extend(['--end', str(end_row)])

        p = subprocess.Popen(
            comando, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            text=True, encoding='utf-8', errors='replace', cwd=BASE_DIR,
            env=os.environ.copy()
        )
        
        # Leitura da saída do processo (sem alterações aqui)
        for linha in iter(p.stdout.readline, ''):
            texto = linha.strip()
            if texto:
                process_output.append(texto)
                # Atualiza contadores com base na saída do script
                if "[SALVO] Processo" in texto: process_status["concluidos"] += 1
                elif "[Acesso ao Processo] Falha" in texto or "[ERRO] Falha ao salvar" in texto or "[ERRO] Processo não cadastrado" in texto:
                    process_status["falhados"] += 1
        p.stdout.close()
        p.wait()
        
        process_output.append("[INFO] Fim da execução do lote.")
        
    except Exception as e:
        process_output.append(f"[ERRO CRÍTICO] Falha ao executar o script RPA: {e}")
    finally:
        process_running = False

# --- FUNÇÃO AUXILIAR PARA PROCESSAR DADOS DO DASHBOARD ---
# (Sem alterações aqui)
def processar_dados_dashboard(dados_log):
    total_processos = len(dados_log)
    
    # Lógica de contagem corrigida para ser explícita
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
    arquivo = request.files.get('arquivo')
    if not arquivo or not arquivo.filename.endswith('.xlsx'):
        flash('Nenhum arquivo .xlsx válido foi enviado.')
        return redirect(url_for('index'))

    # Salva o arquivo original
    upload_path = DATA_DIR / 'entrada_original.xlsx'
    arquivo.save(upload_path)
    
    # Cria a cópia 'entrada.xlsx' que o RPA usará (ou sobrescreve se já existir)
    # Isso preserva o arquivo original caso precisemos reler para outros lotes
    try:
        import shutil
        shutil.copyfile(upload_path, DATA_DIR / 'entrada.xlsx')
        flash(f'Arquivo "{secure_filename(arquivo.filename)}" enviado e copiado para "entrada.xlsx".')
    except Exception as e:
         flash(f'Erro ao copiar arquivo para entrada.xlsx: {e}')
         # Considerar não redirecionar se a cópia falhar?
         
    return redirect(url_for('index'))


# Modificado para receber parâmetros 'start' e 'end'
@app.route('/executar')
def executar_rpa():
    if not (DATA_DIR / 'entrada.xlsx').exists():
        flash('Faça o upload do arquivo "entrada.xlsx" primeiro.')
        return redirect(url_for('index'))

    # Pega os parâmetros da URL
    start_row = request.args.get('start')
    end_row = request.args.get('end')

    # Validação simples (pode ser aprimorada)
    try:
        if start_row: int(start_row)
        if end_row: int(end_row)
    except ValueError:
        flash('Linha inicial ou final inválida. Use apenas números.')
        return redirect(url_for('index'))

    if not process_running:
         # Inicia a thread passando os parâmetros de lote
         threading.Thread(target=executar_script, args=(start_row, end_row)).start()
    else:
        # Informa sobre o lote em execução, se houver
        info = current_batch_info
        msg = "Processo já em execução."
        if info.get('start') or info.get('end'):
             msg += f" (Lote atual: Linha {info.get('start') or 'Início'} até {info.get('end') or 'Fim'})"
        flash(msg)

    # A página executar.html não precisa mudar, ela só exibe o stream
    return render_template('executar.html')


@app.route('/stream')
def stream():
    def event_stream():
        import time
        last_index = 0
        while process_running or last_index < len(process_output):
            # Envia logs acumulados
            current_len = len(process_output)
            if last_index < current_len:
                for i in range(last_index, current_len):
                    yield f"data: {json.dumps({'type': 'log', 'message': process_output[i]})}\n\n"
                last_index = current_len

            # Envia status de progresso atualizado
            # Calcula restantes baseado no total DO LOTE ATUAL
            restantes = process_status["total"] - (process_status["concluidos"] + process_status["falhados"])
            progress_data = {'type': 'progress', **process_status, 'restantes': max(0, restantes)}
            yield f"data: {json.dumps(progress_data)}\n\n"

            if not process_running and last_index == len(process_output):
                break # Garante que saia do loop se o processo acabou e todos os logs foram enviados

            time.sleep(0.5) # Pausa para não sobrecarregar
    return Response(event_stream(), mimetype="text/event-stream")


@app.route('/resultado')
def resultado():
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
# (Sem alterações aqui)
def listar_arquivos(diretorio, extensao, prefixo, formato_data):
    arquivos = sorted(diretorio.glob(f'{prefixo}*.{extensao}'), key=os.path.getmtime, reverse=True)
    lista_formatada = []
    for arq in arquivos:
        try:
            ts_str = arq.name.replace(prefixo, '').replace(f'.{extensao}', '')
            dt_obj = datetime.strptime(ts_str, '%Y-%m-%d_%H-%M-%S')
            nome_exibicao = dt_obj.strftime(formato_data)
        except ValueError:
            nome_exibicao = arq.name # Fallback se o nome não seguir o padrão
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
    path = LOGS_DIR / secure_filename(nome_arquivo) # Segurança extra
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
    path = REPORTS_DIR / secure_filename(nome_arquivo) # Segurança extra
    if not path.is_file():
         flash(f"Arquivo de relatório '{nome_arquivo}' não encontrado.")
         return redirect(url_for('historico_relatorios'))
    try:
        df = pd.read_csv(path, sep=';', encoding='utf-8-sig')
        # Limita o número de linhas exibidas para evitar sobrecarga no navegador
        max_rows_display = 500 
        table_html = df.head(max_rows_display).to_html(classes='table-auto w-full text-left', index=False)
        if len(df) > max_rows_display:
             table_html += f"<p class='mt-4 text-sm text-yellow-600'>Aviso: Exibindo apenas as primeiras {max_rows_display} linhas de {len(df)}. Baixe o CSV para ver o relatório completo.</p>"
        return render_template('ver_relatorio_csv.html', nome_arquivo=nome_arquivo, tabela=table_html)
    except Exception as e:
        flash(f"Erro ao ler o relatório '{nome_arquivo}': {e}")
        return redirect(url_for('historico_relatorios'))


@app.route('/baixar_log/<path:nome>')
def baixar_log(nome):
    return send_from_directory(LOGS_DIR, secure_filename(nome), as_attachment=True)

@app.route('/baixar_relatorio_csv/<path:nome>')
def baixar_relatorio_csv(nome):
    return send_from_directory(REPORTS_DIR, secure_filename(nome), as_attachment=True)

if __name__ == '__main__':
    # Usar host='0.0.0.0' torna acessível na rede local, útil para testes
    app.run(debug=True, host='0.0.0.0', port=8080)