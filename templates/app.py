from flask import Flask, render_template, request, redirect, url_for, send_from_directory, flash, Response
from werkzeug.utils import secure_filename
from pathlib import Path
from datetime import datetime
import subprocess
import os
import threading
import json

app = Flask(__name__)
app.secret_key = 'secreta_onesync'

# --- ESTRUTURA DE DIRETÓRIOS ---
BASE_DIR = Path(__file__).resolve().parent
UPLOAD_FOLDER = BASE_DIR / 'uploads'
DATA_DIR = BASE_DIR / 'data'
LOGS_DIR = DATA_DIR / 'logs_execucao' # 👈 CORRIGIDO: Aponta para a pasta correta

# Garante que as pastas existem
UPLOAD_FOLDER.mkdir(parents=True, exist_ok=True)
LOGS_DIR.mkdir(parents=True, exist_ok=True)

# --- VARIÁVEIS GLOBAIS PARA EXECUÇÃO EM TEMPO REAL ---
process_output = []
process_running = False

# --- FUNÇÃO DE EXECUÇÃO DO RPA EM BACKGROUND ---
def executar_script():
    """Executa o script RPA em uma thread separada para não bloquear a interface."""
    global process_output, process_running
    process_output = []
    process_running = True
    try:
        print("[LOG] Iniciando subprocesso do RPA...", flush=True)

        # Comando para executar o script Python
        comando = [
            'python', '-u', str(BASE_DIR / 'rpa_atualizar_pasta.py')
        ]

        # Inicia o processo
        p = subprocess.Popen(
            comando,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding='utf-8',
            errors='replace',
            cwd=BASE_DIR # Garante que o script execute no diretório correto
        )

        # Captura a saída em tempo real
        for linha in iter(p.stdout.readline, ''):
            texto = linha.strip()
            if texto:
                print(f"[RPA] {texto}", flush=True)
                process_output.append(texto)

        p.stdout.close()
        return_code = p.wait()

        if return_code != 0:
            process_output.append(f"[ERRO] O script RPA finalizou com código de erro: {return_code}")

    except Exception as e:
        erro = f"[ERRO CRÍTICO] Falha ao executar o script RPA: {str(e)}"
        print(erro, flush=True)
        process_output.append(erro)
    finally:
        process_running = False
        print("[LOG] Subprocesso do RPA finalizado.", flush=True)


# --- ROTAS DA APLICAÇÃO WEB ---

@app.route('/')
def index():
    """Página inicial para upload de arquivo."""
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    """Recebe o arquivo .xlsx e salva no local correto."""
    arquivo = request.files.get('arquivo')
    if not arquivo or arquivo.filename == '':
        flash('Nenhum arquivo enviado.')
        return redirect(url_for('index'))

    if not arquivo.filename.endswith('.xlsx'):
        flash('Apenas arquivos .xlsx são aceitos.')
        return redirect(url_for('index'))

    # Define o nome do arquivo padrão para ser lido pelo RPA
    caminho_destino = DATA_DIR / 'entrada.xlsx'
    arquivo.save(caminho_destino)

    flash(f'Arquivo "{secure_filename(arquivo.filename)}" enviado com sucesso e pronto para execução.')
    return redirect(url_for('index'))

@app.route('/executar')
def executar_rpa():
    """Página que mostra o log de execução em tempo real."""
    global process_running
    if not (DATA_DIR / 'entrada.xlsx').exists():
        flash('Nenhum arquivo "entrada.xlsx" encontrado. Por favor, faça o upload primeiro.')
        return redirect(url_for('index'))

    if not process_running:
        thread = threading.Thread(target=executar_script)
        thread.start()
    return render_template('executar.html')

@app.route('/stream')
def stream():
    """Fornece o fluxo de dados (log) para a página de execução."""
    def event_stream():
        import time
        last_index = 0
        while process_running or last_index < len(process_output):
            while last_index < len(process_output):
                data = process_output[last_index]
                yield f"data: {data}\n\n"
                last_index += 1
            time.sleep(0.1)
    return Response(event_stream(), mimetype="text/event-stream")


@app.route('/resultado')
def resultado():
    """Exibe um dashboard de análise da última execução do RPA."""
    try:
        # 1. Encontrar o arquivo de log mais recente
        logs = sorted(LOGS_DIR.glob('log_*.json'), key=os.path.getmtime, reverse=True)
        if not logs:
            return render_template('resultado.html', dashboard_data=None)

        ultimo_log_path = logs[0]
        with open(ultimo_log_path, 'r', encoding='utf-8') as f:
            dados_log = json.load(f)

        # 2. Calcular as Métricas de KPI
        total_processos = len(dados_log)
        sucessos = 0
        falhas = 0
        tempo_total_execucao = 0.0

        for processo in dados_log:
            if processo.get("Status Geral") == "Sucesso":
                sucessos += 1
            else:
                falhas += 1
            for duracao in processo.get("Duracao por etapa (s)", {}).values():
                tempo_total_execucao += duracao
        
        # ================================================================= #
        # ✨ INÍCIO DO NOVO CÓDIGO: ANÁLISE PARA GRÁFICOS ✨
        # ================================================================= #
        
        # Gráfico 1: Duração por Etapa
        duracao_por_etapa = {}
        for processo in dados_log:
            for etapa, duracao in processo.get("Duracao por etapa (s)", {}).items():
                duracao_por_etapa[etapa] = duracao_por_etapa.get(etapa, 0) + duracao
        
        # Ordena as etapas pela duração (do maior para o menor) para um gráfico mais claro
        etapas_ordenadas = sorted(duracao_por_etapa.items(), key=lambda item: item[1], reverse=True)
        
        # Prepara as listas de labels (nomes das etapas) e dados (tempos) para o Chart.js
        grafico_duracao_labels = [etapa for etapa, duracao in etapas_ordenadas]
        grafico_duracao_data = [round(duracao, 2) for etapa, duracao in etapas_ordenadas]
        
        # ================================================================= #
        # ✨ FIM DO NOVO CÓDIGO ✨
        # ================================================================= #

        # 3. Preparar todos os dados para enviar ao template
        dashboard_data = {
            'nome_arquivo': ultimo_log_path.name,
            'total_processos': total_processos,
            'sucessos': sucessos,
            'falhas': falhas,
            'percentual_sucesso': (sucessos / total_processos * 100) if total_processos > 0 else 0,
            'percentual_falha': (falhas / total_processos * 100) if total_processos > 0 else 0,
            'tempo_total_execucao': tempo_total_execucao,
            
            # Adiciona os dados do gráfico de duração
            'grafico_duracao': {
                'labels': grafico_duracao_labels,
                'data': grafico_duracao_data
            },

            'dados_brutos': dados_log 
        }

        return render_template('resultado.html', dashboard_data=dashboard_data)

    except Exception as e:
        flash(f"Ocorreu um erro ao gerar o dashboard de resultados: {e}")
        return redirect(url_for('index'))


@app.route('/historico')
def historico():
    """Lista todos os relatórios de log com nomes formatados."""
    try:
        # Busca por arquivos .json e ordena do mais novo para o mais antigo
        arquivos_log = sorted(
            LOGS_DIR.glob('log_*.json'), 
            key=os.path.getmtime, 
            reverse=True
        )

        relatorios_formatados = []
        for arquivo in arquivos_log:
            nome_original = arquivo.name
            nome_exibicao = nome_original # Valor padrão caso a formatação falhe

            try:
                # Extrai a string de timestamp do nome do arquivo
                # Ex: de "log_2025-07-16_11-59-28.json" para "2025-07-16_11-59-28"
                timestamp_str = nome_original.replace('log_', '').replace('.json', '')
                
                # Converte a string em um objeto de data e hora
                dt_obj = datetime.strptime(timestamp_str, '%Y-%m-%d_%H-%M-%S')
                
                # Formata a data e hora para um formato amigável
                nome_exibicao = dt_obj.strftime('Execução de %d/%m/%Y às %H:%M:%S')

            except ValueError:
                # Se o nome do arquivo não corresponder ao padrão, ele será exibido como está
                print(f"Aviso: Não foi possível formatar o nome do arquivo '{nome_original}'.")

            relatorios_formatados.append({
                'original': nome_original,
                'exibicao': nome_exibicao
            })
        
        # Passa a lista formatada para o template
        return render_template('historico.html', arquivos=relatorios_formatados)

    except Exception as e:
        flash(f"Erro ao acessar o histórico: {e}")
        return render_template('historico.html', arquivos=[])


@app.route('/relatorio/<path:nome_arquivo>')
def ver_relatorio(nome_arquivo):
    """Exibe o conteúdo de um relatório de log específico."""
    try:
        caminho_arquivo = LOGS_DIR / nome_arquivo
        if not caminho_arquivo.exists():
            flash(f"Relatório '{nome_arquivo}' não encontrado.")
            return redirect(url_for('historico'))

        with open(caminho_arquivo, 'r', encoding='utf-8') as f:
            # Carrega o conteúdo JSON do arquivo
            dados_json = json.load(f)

        # Formata o JSON para uma string bonita e legível
        relatorio_formatado = json.dumps(dados_json, indent=2, ensure_ascii=False)
        
        # Reutiliza o template 'resultado.html' para exibir o conteúdo
        return render_template('resultado.html', relatorio=relatorio_formatado, nome_arquivo=nome_arquivo)

    except Exception as e:
        flash(f"Erro ao ler o relatório '{nome_arquivo}': {e}")
        return redirect(url_for('historico'))

@app.route('/baixar_relatorio/<path:nome>')
def baixar_relatorio(nome):
    """Permite o download de um arquivo de log específico."""
    return send_from_directory(LOGS_DIR, nome, as_attachment=True)


if __name__ == '__main__':
    app.run(debug=True, port=8080)