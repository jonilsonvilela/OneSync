from playwright.sync_api import sync_playwright
from pathlib import Path
from datetime import datetime
import pandas as pd
import json
import os
import time
import sys

# Altera a codificação APENAS se estiver rodando em um terminal interativo no Windows
if sys.stdout.isatty() and os.name == 'nt':
    sys.stdout.reconfigure(encoding='utf-8')

MAPA_ESCRITORIO_RESPONSAVEL = {

    "MDR Advocacia / Área operacional / Ativos / Autor": 26,
    "MDR Advocacia / Área operacional / Ativos / Réu": 27,
    "MDR Advocacia / Área operacional / Ativos / Trabalhista": 37,
    "MDR Advocacia / Área operacional / Ativos / Administrativo": 38,
    "MDR Advocacia / Área operacional / Banco do Brasil / Autor": 22,
    "MDR Advocacia / Área operacional / Banco do Brasil / Réu": 23,
    "MDR Advocacia / Área operacional / Banco do Brasil / Trabalhista": 24,
    "MDR Advocacia / Área operacional / Banco do Brasil / Interessado": 40,
    "MDR Advocacia / Área operacional / Banese / Autor": 42,
    "MDR Advocacia / Área operacional / Banese / Réu": 44,
    "MDR Advocacia / Área operacional / Bradesco / Réu": 45,
    "MDR Advocacia / Área operacional / Bradesco / Autor": 46,
}

MAPA_NEGOCIACAO_HONORARIO = {
    'Lote 12 - Rio Grande do Norte / Autor / Hon - 0000001/001': 1,
    'Lote 12 - Rio Grande do Norte / Réu / Hon - 0000001/002': 2,
    'Lote 12 - Rio Grande do Norte / Trabalhista / Hon - 0000001/003': 3,
    'Lote 1 - Acre e Rondônia / Autor / Hon - 0000001/004': 4,
    'Lote 1 - Acre e Rondônia / Réu / Hon - 0000001/005': 5,
    'Lote 1 - Acre e Rondônia / Trabalhista / Hon - 0000001/006': 6,
    'Lote 2 - Amazonas e Roraima / Autor / Hon - 0000001/007': 7,
    'Lote 2 - Amazonas e Roraima / Réu / Hon - 0000001/008': 8,
    'Lote 2 - Amazonas e Roraima / Trabalhista / Hon - 0000001/009': 9,
    'Lote 3 - Amapá e Pará / Autor / Hon - 0000001/010': 10,
    'Lote 3 - Amapá e Pará / Réu / Hon - 0000001/011': 11,
    'Lote 3 - Amapá e Pará / Trabalhista / Hon - 0000001/012': 12,
    'Juizado Especial - Polo Passivo / Ativos / Hon - 0000002/001': 13,
    'Cível Comum - Polo Passivo / Ativos / Hon - 0000002/002': 14,
    'Trabalhista / Ativos / Hon - 0000002/003': 15,
    'Polo Ativo - Cível Comum / Ativos / Hon - 0000002/004': 16,
    'Administrativo / Ativos / Hon - 0000002/005': 17,
    'Negociação Padrão / Bradesco / Hon - 0000003/001': 18,
    'Negociação Padrão / Banese / Hon - 0000004/001': 19,
}

MAPA_CENTRO_CUSTO = {
    "MDR - MANHATTAN": 39,
    "MDR Advocacia / Área Administrativa / Área Operacional / Ativos / Autor": 26,
    "MDR Advocacia / Área Administrativa / Área Operacional / Ativos / Réu": 27,
    "MDR Advocacia / Área Administrativa / Área Operacional / Ativos / Trabalhista": 37,
    "MDR Advocacia / Área Administrativa / Área Operacional / Ativos / Administrativo": 38,
    "MDR Advocacia / Área Administrativa / Área Operacional / Banco do Brasil / Autor": 22,
    "MDR Advocacia / Área Administrativa / Área Operacional / Banco do Brasil / Réu": 23,
    "MDR Advocacia / Área Administrativa / Área Operacional / Banco do Brasil / Trabalhista": 24,
    "MDR Advocacia / Área Administrativa / Área Operacional / Banco do Brasil / Interessado": 40,
    "MDR Advocacia / Área Administrativa / Área Operacional / Banese / Autor": 42,
    "MDR Advocacia / Área Administrativa / Área Operacional / Banese / Réu": 44,
    "MDR Advocacia / Área Administrativa / Área Operacional / Bradesco / Réu": 45,
    "MDR Advocacia / Área Administrativa / Área Operacional / Bradesco / Autor": 46,
}

def iniciar_log_processo(numero_processo: str):
    return {
        "Processo": numero_processo,
        "Status Geral": "Em andamento",
        "Erros": [],
        "Sucessos": [],
        "Duracao por etapa (s)": {}
    }

def adicionar_evento(log: dict, etapa: str, tipo: str, item: str, mensagem: str = "", duracao: float = None):
    entrada = {"etapa": etapa, "item": item}
    if mensagem:
        entrada["mensagem"] = mensagem

    if tipo == "erro":
        log["Erros"].append(entrada)
    elif tipo == "sucesso":
        log["Sucessos"].append(entrada)

    if duracao is not None:
        log.setdefault("Duracao por etapa (s)", {})[etapa] = duracao

def salvar_log_execucao(lista_logs: list):
    ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    pasta = Path(__file__).resolve().parent / "data" / "logs_execucao"
    pasta.mkdir(parents=True, exist_ok=True)

    path = pasta / f"log_{ts}.json"

    # Atualiza status geral
    for log in lista_logs:
        if any("mensagem" in erro for erro in log.get("Erros", [])):
            log["Status Geral"] = "Falha"
        else:
            log["Status Geral"] = "Sucesso"

    with open(path, "w", encoding="utf-8") as f:
        json.dump(lista_logs, f, ensure_ascii=False, indent=2)

    print(f"[🧾] Log da execução salvo em: {path}")

def login_legalone(page, usuario, senha):
    """LOGIN NO PORTAL LEGAL ONE"""
    inicio = time.time()
    etapa = "Login"
    try:
        page.goto("https://signon.thomsonreuters.com/?productId=L1NJ&returnto=https%3a%2f%2flogin.novajus.com.br%2fOnePass%2fLoginOnePass%2f")
        page.fill('input#Username', usuario)
        page.fill('input#Password', senha)
        page.click('button[name="SignIn"]', no_wait_after=True)
        page.wait_for_timeout(10000)

        return {
            "etapa": etapa,
            "duracao": round(time.time() - inicio, 2),
            "status": "Sucesso"
        }

    except Exception as e:
        return {
            "etapa": etapa,
            "duracao": round(time.time() - inicio, 2),
            "status": "Falha",
            "mensagem": str(e)
        }

def acessar_processo_para_edicao(page, numero_processo):
    """ACESSO AO PROCESSO PARA EDIÇÃO"""
    inicio = time.time()
    etapa = "Acesso ao Processo"

    try:
        page.goto("https://mdradvocacia.novajus.com.br/processos/processos/search")
        campo = page.locator('input#Search')
        campo.fill(numero_processo)
        campo.press("Enter")
        page.wait_for_timeout(3000)

        if page.query_selector('p.legalone-grid-counter.result-header[data-count="0"]'):
            raise ValueError("Processo não cadastrado")

        page.wait_for_selector("table.webgrid", timeout=30000)
        page.locator("span.grid-overflow-icon").click()
        page.locator("a.grid-edit-action-row").click()
        page.wait_for_selector('button[name="ButtonSave"]', timeout=30000)

        return {
            "etapa": etapa,
            "duracao": round(time.time() - inicio, 2),
            "status": "Sucesso"
        }

    except Exception as e:
        return {
            "etapa": etapa,
            "duracao": round(time.time() - inicio, 2),
            "status": "Falha",
            "mensagem": str(e)
        }

def expandir_paineis(page):
    """EXPANSÃO DOS PAINÉIS DE EDIÇÃO"""
    inicio = time.time()
    etapa = "Expandir Painéis"

    try:
        paineis = page.locator('p.panel-title')
        total = paineis.count()
    except Exception as e:
        return {
            "etapa": etapa,
            "duracao": round(time.time() - inicio, 2),
            "status": "Falha",
            "mensagem": f"Erro ao localizar painéis: {e}"
        }

    for i in range(total):
        try:
            icone = paineis.nth(i).locator('span.arrow')
            if "arrow-expanded" not in (icone.get_attribute("class") or ""):
                paineis.nth(i).click()
                time.sleep(0.2)
        except:
            continue

    return {
        "etapa": etapa,
        "duracao": round(time.time() - inicio, 2),
        "status": "Sucesso"
    }

def preencher_lookups_em_cascata(page, dados_linha, mapeamento):
    """PREENCHIMENTO DOS CAMPOS EM CASCATA"""
    inicio = time.time()
    etapa = "Cascata"
    mensagens_erro = []
    campos_preenchidos = 0

    print(" ├─ Preenchendo campos em cascata (UF → Cidade → Justiça → Instância → Classe → Assunto → Comarca/Foro → Número da Vara → Tipo da Vara)...")

    ordem = [
        "Estado (UF)",
        "Cidade",
        "Justiça (CNJ)",
        "Instância (CNJ)",
        "Classe (CNJ)",
        "Assunto (CNJ)",
        "Comarca/Foro",
        "Número da Vara",
        "Tipo da Vara"
    ]

    for atual in ordem:
        try:
            valor = dados_linha.get(atual)
            if valor is None or str(valor).strip() == "":
                print(f"    ⚠️ Campo '{atual}' sem valor preenchido na planilha, pulando.")
                continue

            campo = next((c for c in mapeamento if c.get("descricao") == atual), None)
            if not campo or not campo.get("id"):
                print(f"    ⚠️ Campo '{atual}' não encontrado no mapeamento.")
                continue

            # Caso especial: "Número da Vara" é um campo simples
            if atual == "Número da Vara":
                seletor = f'input[id="{campo["id"]}"], input[id*="{campo["id"]}"]'
                try:
                    page.wait_for_selector(seletor, timeout=10000)
                    el_input = page.query_selector(seletor)
                    if not el_input or not el_input.is_enabled():
                        print(f"    ⚠️ Campo '{atual}' não localizado ou desabilitado.")
                        continue

                    el_input.fill(str(valor))
                    el_input.press("Tab")
                    page.wait_for_timeout(300)
                    print(f"    ✔ Campo '{atual}' preenchido com sucesso.")
                    campos_preenchidos += 1
                except Exception as e:
                    print(f"    ⚠️ Erro ao preencher '{atual}': {e}")
                    mensagens_erro.append(f"{atual}: {e}")
                continue

            # Campos com ID dinâmico (Assunto)
            if atual == "Assunto (CNJ)":
                identificador = campo["id"].split("__")[-1]
                seletor = f'input[id^="Assuntos_"][id$="__{identificador}"]:not([type="hidden"])'
            else:
                seletor = f'input.search.ac_input[id="{campo["id"]}"], input.search.ac_input[id*="{campo["id"]}"]'

            # Aguarda desbloqueio do campo com lógica padrão para lookup
            try:
                page.wait_for_function(
                    """(selector) => {
                        const el = document.querySelector(selector);
                        if (!el) return false;
                        const wrapper = el.closest('.lookup');
                        return !el.readOnly && wrapper && !wrapper.classList.contains('disabled');
                    }""",
                    arg=seletor,
                    timeout=20000
                )
                print(f"    ⏳ Campo '{atual}' agora disponível.")
                page.wait_for_timeout(500)
            except:
                print(f"    ⚠️ Campo '{atual}' não desbloqueou a tempo.")
                mensagens_erro.append(f"{atual}: timeout de desbloqueio")
                continue

            el_input = page.query_selector(seletor)
            if not el_input:
                print(f"    ⚠️ Campo '{atual}' não localizado.")
                mensagens_erro.append(f"{atual}: não localizado")
                continue
            if not el_input.is_enabled():
                print(f"    ⚠️ Campo '{atual}' está desabilitado.")
                mensagens_erro.append(f"{atual}: desabilitado")
                continue

            print(f"    🖱 Preenchendo campo '{atual}' com valor: '{valor}'")

            el_input.click()
            el_input.fill(str(valor))
            page.wait_for_timeout(500)

            el_input.press("Enter")
            page.wait_for_timeout(1000)
            el_input.press("ArrowDown")
            page.wait_for_timeout(300)
            el_input.press("Enter")
            page.wait_for_timeout(300)
            el_input.press("Tab")
            page.wait_for_timeout(300)

            print(f"    ✔ {atual} preenchido com sucesso.")
            campos_preenchidos += 1

        except Exception as e:
            print(f"    ⚠️ Erro ao preencher '{atual}': {e}")
            mensagens_erro.append(f"{atual}: {e}")

    if campos_preenchidos > 0:
        return {
            "etapa": etapa,
            "duracao": round(time.time() - inicio, 2),
            "status": "Sucesso"
        }
    else:
        return {
            "etapa": etapa,
            "duracao": round(time.time() - inicio, 2),
            "status": "Falha",
            "mensagem": "; ".join(mensagens_erro)
        }

def preencher_outros_envolvidos(page, dados_linha):
    """PREENCHIMENTO DOS ENVOLVIDOS"""
    import time
    inicio = time.time()
    etapa = "Envolvidos"
    erros = []
    sucessos = 0

    print(" ├─ Preenchendo Envolvidos...")

    situacoes = [s.strip() for s in str(dados_linha.get("Situação do Envolvido", "")).split(";") if s.strip()]
    posicoes  = [s.strip() for s in str(dados_linha.get("Posição Envolvido", "")).split(";") if s.strip()]
    nomes     = [s.strip() for s in str(dados_linha.get("Envolvido", "")).split(";") if s.strip()]
    tipos     = [s.strip() for s in str(dados_linha.get("Tipo do Envolvido", "")).split(";")] if "Tipo do Envolvido" in dados_linha else []
    docs      = [s.strip() for s in str(dados_linha.get("CPF/CNPJ do Envolvido", "")).split(";")] if "CPF/CNPJ do Envolvido" in dados_linha else []

    total = max(len(situacoes), len(posicoes), len(nomes))
    if total == 0:
        return {
            "etapa": etapa,
            "duracao": round(time.time() - inicio, 2),
            "status": "Sucesso"  # Nenhum para preencher, mas não é erro
        }

    try:
        for _ in range(10):
            envolvidos = page.locator("ul.outros-envolvidos-list > li")
            if envolvidos.count() > 0:
                break
            page.wait_for_timeout(300)

        quantidade_existente = envolvidos.count()
        faltam_adicionar = total - quantidade_existente

        for i in range(faltam_adicionar):
            try:
                botao = page.locator('a#add_outro_envolvido')
                botao.scroll_into_view_if_needed()
                botao.wait_for(state="visible", timeout=5000)
                botao.click()
                page.wait_for_timeout(800)
            except Exception as e:
                msg = f"Erro ao adicionar envolvido {quantidade_existente + i + 1}: {e}"
                print(f"[⚠️] {msg}")
                erros.append(msg)

        for _ in range(10):
            envolvidos = page.locator("ul.outros-envolvidos-list > li")
            if envolvidos.count() >= total:
                break
            page.wait_for_timeout(300)
        else:
            msg = f"Apenas {envolvidos.count()} campos de envolvido visíveis, esperado {total}"
            print(f"[ERRO] {msg}")
            return {
                "etapa": etapa,
                "duracao": round(time.time() - inicio, 2),
                "status": "Falha",
                "mensagem": msg
            }

        for i in range(total):
            grupo = envolvidos.nth(i)

            if i < len(situacoes):
                try:
                    grupo.locator('select[id*="SituacaoEnvolvidoId"]').select_option(label=situacoes[i])
                    print(f"  │   ├─ Situação {i+1}: {situacoes[i]}")
                except Exception as e:
                    erros.append(f"Situação {i+1}: {e}")

            if i < len(posicoes):
                try:
                    pos_input = grupo.locator('input[id*="PosicaoEnvolvidoText"]')
                    if pos_input and pos_input.is_visible():
                        pos_input.first.fill(posicoes[i])
                        page.wait_for_timeout(300)
                        sucesso = preencher_lookup_com_validacao(page, pos_input.first, posicoes[i])
                        if sucesso:
                            print(f"  │   ├─ Posição {i+1}: {posicoes[i]}")
                        else:
                            erros.append(f"Validação falhou para posição {i+1}")
                    else:
                        erros.append(f"Campo Posição {i+1} não visível")
                except Exception as e:
                    erros.append(f"Erro posição {i+1}: {e}")

            if i < len(nomes):
                try:
                    nome = nomes[i]
                    tipo = tipos[i] if i < len(tipos) else "Pessoa física"
                    doc  = docs[i] if i < len(docs) else ""

                    bloco_lookup = grupo.locator('div[id*="lookup_envolvido"]')
                    input_nome = bloco_lookup.locator('input[id*="EnvolvidoText"]')

                    if not input_nome or not input_nome.is_visible():
                        erros.append(f"Campo Nome {i+1} não visível")
                        continue

                    input_nome.first.click()
                    input_nome.first.fill(nome)
                    page.wait_for_timeout(300)
                    input_nome.first.press("Enter")
                    page.wait_for_timeout(300)
                    input_nome.first.press("ArrowDown")
                    page.wait_for_timeout(500)

                    if page.query_selector("div.empty-box"):
                        print(f"  ⚠️ Envolvido {i+1} não encontrado. Abrindo modal...")
                        botao_plus = bloco_lookup.locator('.lookup-button.lookup-new')
                        if botao_plus:
                            botao_plus.click()
                            page.wait_for_selector('form#contatoForm', timeout=10000)
                            sucesso_modal = preencher_modal_adverso(page, nome, tipo, doc)
                            if sucesso_modal:
                                print(f"  ✅ Envolvido {i+1} cadastrado com sucesso.")
                                sucessos += 1
                            else:
                                erros.append(f"Falha ao cadastrar Envolvido {i+1}")
                        else:
                            erros.append(f"Botão '+' do Envolvido {i+1} não encontrado")
                    else:
                        input_nome.first.press("Enter")
                        page.wait_for_timeout(300)
                        input_nome.first.press("Tab")
                        print(f"  │   ├─ Nome {i+1}: {nome}")
                        sucessos += 1

                except Exception as e:
                    erros.append(f"Erro Nome {i+1}: {e}")

    except Exception as e:
        erros.append(f"Erro geral em Envolvidos: {e}")

    if sucessos > 0:
        return {
            "etapa": etapa,
            "duracao": round(time.time() - inicio, 2),
            "status": "Sucesso"
        }
    else:
        return {
            "etapa": etapa,
            "duracao": round(time.time() - inicio, 2),
            "status": "Falha",
            "mensagem": "; ".join(erros)
        }

def preencher_objetos(page, dados_linha):
    """PREENCHIMENTO DOS OBJETOS"""
    import time
    inicio = time.time()
    etapa = "Objetos"
    erros = []
    sucessos = 0

    print("  ├─ Preenchendo Objetos...")
    nomes = [s.strip() for s in str(dados_linha.get("Nome do Objeto", "")).split(";") if s.strip()]
    observacoes = [s.strip() for s in str(dados_linha.get("Observações do objeto", "")).split(";") if s.strip()]

    total = max(len(nomes), len(observacoes))
    if total == 0:
        return {
            "etapa": etapa,
            "duracao": round(time.time() - inicio, 2),
            "status": "Sucesso"
        }

    try:
        for _ in range(10):
            objetos = page.locator("ul.objetos-list > li")
            if objetos.count() > 0:
                break
            page.wait_for_timeout(300)

        quantidade_existente = objetos.count()
        faltam_adicionar = total - quantidade_existente

        for i in range(faltam_adicionar):
            try:
                botao = page.locator('a#add_objetos')
                botao.scroll_into_view_if_needed()
                botao.wait_for(state="visible", timeout=5000)
                botao.click()
                page.wait_for_timeout(800)
            except Exception as e:
                msg = f"Erro ao adicionar objeto {quantidade_existente + i + 1}: {e}"
                print(f"[⚠️] {msg}")
                erros.append(msg)

        for _ in range(10):
            objetos = page.locator("ul.objetos-list > li")
            if objetos.count() >= total:
                break
            page.wait_for_timeout(300)
        else:
            msg = f"Apenas {objetos.count()} campos de objeto visíveis, esperado {total}"
            print(f"[ERRO] {msg}")
            return {
                "etapa": etapa,
                "duracao": round(time.time() - inicio, 2),
                "status": "Falha",
                "mensagem": msg
            }

        for i in range(total):
            grupo = objetos.nth(i)

            if i < len(nomes):
                try:
                    input_nome = grupo.locator('input[id*="ObjetoText"]')
                    sucesso = preencher_lookup_com_validacao(page, input_nome, nomes[i])
                    if sucesso:
                        print(f"  │   ├─ Nome do Objeto {i+1}: {nomes[i]}")
                        sucessos += 1
                    else:
                        msg = f"Validação falhou para Nome do Objeto {i+1}"
                        print(f"[⚠️] {msg}")
                        erros.append(msg)
                except Exception as e:
                    erros.append(f"Erro Nome {i+1}: {e}")

            if i < len(observacoes):
                try:
                    grupo.locator('textarea[id*="Observacoes"]').fill(observacoes[i])
                    print(f"  │   ├─ Observações do Objeto {i+1}")
                except Exception as e:
                    erros.append(f"Erro Observações {i+1}: {e}")

    except Exception as e:
        erros.append(f"Erro geral ao preencher objetos: {e}")

    if sucessos > 0:
        return {
            "etapa": etapa,
            "duracao": round(time.time() - inicio, 2),
            "status": "Sucesso"
        }
    else:
        return {
            "etapa": etapa,
            "duracao": round(time.time() - inicio, 2),
            "status": "Falha",
            "mensagem": "; ".join(erros)
        }

def preencher_pedidos(page, dados_linha, mapeamento):
    """PREENCHIMENTO DOS PEDIDOS"""
    import time
    import pandas as pd

    inicio = time.time()
    etapa = "Pedidos"
    erros = []
    pedidos_sucesso = 0

    def aguardar_blocos_pedidos(page, quantidade_esperada, timeout_ms=5000):
        inicio_espera = time.time()
        while time.time() - inicio_espera < (timeout_ms / 1000):
            visiveis = page.locator("ul.pedidos-list > li:visible").count()
            if visiveis >= quantidade_esperada:
                return True
            page.wait_for_timeout(300)
        print(f"[⚠️] Só {visiveis} blocos visíveis após {timeout_ms}ms (esperado: {quantidade_esperada})")
        return False

    print(" ├─ Preenchendo Pedidos...")
    campos_pedidos = [
        "Tipo de Pedido", "Contingência", "Data do Pedido", "Data do Julgamento",
        "Probabilidade de Êxito", "Situação do Pedido", "Valor do Pedido",
        "Valor da Condenação", "Observação do Pedido"
    ]

    if not any(dados_linha.get(campo) for campo in campos_pedidos):
        return {
            "etapa": etapa,
            "duracao": round(time.time() - inicio, 2),
            "status": "Sucesso"
        }

    valores_por_campo = {
        campo: [v.strip() for v in str(dados_linha.get(campo)).split(";")] if pd.notna(dados_linha.get(campo)) and str(dados_linha.get(campo)).strip() != "" else []
        for campo in campos_pedidos
    }

    num_pedidos = max(len(v) for v in valores_por_campo.values())
    print(f"  ├─ Total detectado: {num_pedidos} pedido(s)")

    if not page.query_selector("ul.pedidos-list"):
        msg = "Lista de pedidos não localizada."
        print(f"[ERRO] {msg}")
        return {
            "etapa": etapa,
            "duracao": round(time.time() - inicio, 2),
            "status": "Falha",
            "mensagem": msg
        }

    while page.locator("ul.pedidos-list > li:visible").count() < num_pedidos:
        try:
            botao = page.locator('a#add_pedido')
            botao.scroll_into_view_if_needed()
            botao.wait_for(state="visible", timeout=5000)
            botao.click()
            page.wait_for_timeout(800)
        except Exception as e:
            erros.append(f"Erro ao adicionar pedido: {e}")
            break

    aguardar_blocos_pedidos(page, num_pedidos, timeout_ms=5000)
    pedidos_li = page.locator("ul.pedidos-list > li:visible")
    print(f"  ├─ {pedidos_li.count()} blocos de pedido visíveis")

    campos_com_lookup = {
        "Tipo de Pedido", "Contingência", "Situação do Pedido", "Probabilidade de Êxito"
    }

    for i in range(num_pedidos):
        li = pedidos_li.nth(i)
        try:
            for campo in campos_pedidos:
                valores = valores_por_campo.get(campo, [])
                if i >= len(valores):
                    continue
                valor = valores[i]
                if not valor.strip():
                    continue

                campo_json = next((c for c in mapeamento if c.get("descricao") == campo), None)
                if campo_json is None:
                    continue

                identificador = campo_json["id"].split("__")[-1]
                seletor = f'input[id^="Pedidos_"][id$="__{identificador}"]:not([type="hidden"]), ' \
                          f'select[id^="Pedidos_"][id$="__{identificador}"], ' \
                          f'textarea[id^="Pedidos_"][id$="__{identificador}"]'

                campos_possiveis = li.locator(seletor)
                if campos_possiveis.count() == 0:
                    continue

                campo_alvo = campos_possiveis.first
                campo_alvo.wait_for(state="visible", timeout=5000)
                page.wait_for_timeout(300)

                if not campo_alvo.is_enabled():
                    continue

                tag = campo_alvo.evaluate("e => e.tagName.toLowerCase()")

                if campo in campos_com_lookup:
                    if campo == "Probabilidade de Êxito":
                        try:
                            campo_alvo.select_option(label="Êxito")
                        except:
                            pass
                        try:
                            input_text = li.locator('input[id*="__ProbabilidadeText"]')
                            input_text.wait_for(state="visible", timeout=3000)
                            preencher_lookup_com_validacao(page, input_text, valor)
                        except:
                            pass
                    else:
                        preencher_lookup_com_validacao(page, campo_alvo, valor)

                elif tag == "input":
                    if campo == "Data do Julgamento":
                        campo_alvo.fill(valor)
                        campo_alvo.press("Tab")
                        print(f"  🗓️ Data do Julgamento preenchida com: {valor}")
                        continue

                    elif "Data" in campo:
                        campo_alvo.click()
                        campo_alvo.press("Control+A")
                        campo_alvo.press("Backspace")
                        campo_alvo.type(valor, delay=50)
                        campo_alvo.press("Tab")
                    else:
                        campo_alvo.fill(valor)
                        campo_alvo.press("Tab")

                elif tag == "select":
                    campo_alvo.select_option(label=valor)
                    page.wait_for_timeout(300)

                elif tag == "textarea":
                    campo_alvo.fill(valor)

            pedidos_sucesso += 1
            print(f"  │   ✔ Pedido {i+1} preenchido com sucesso.")

        except Exception as e:
            erros.append(f"Pedido {i+1} → erro: {e}")

    if pedidos_sucesso > 0:
        return {
            "etapa": etapa,
            "duracao": round(time.time() - inicio, 2),
            "status": "Sucesso"
        }
    else:
        return {
            "etapa": etapa,
            "duracao": round(time.time() - inicio, 2),
            "status": "Falha",
            "mensagem": "; ".join(erros)
        }

def atualizar_campos_usando_mapeamento(page, dados_linha, mapeamento):
    """PREENCHIMENTO DOS CAMPOS GERAIS"""
    import time
    inicio = time.time()
    etapa = "Campos Gerais"
    campos_sucesso = []
    campos_falha = []

    print(" ├─ Preenchendo campos gerais...")

    campos_ja_tratados = {
        "Processo",
        "Estado (UF)", "Cidade", "Justiça (CNJ)", "Instância (CNJ)",
        "Classe (CNJ)", "Assunto (CNJ)",
        "Situação do Envolvido", "Posição Envolvido", "Envolvido",
        "Nome do Objeto", "Observações do objeto",
        "Tipo de Pedido", "Contingência", "Data do Pedido", "Data do Julgamento",
        "Probabilidade de Êxito", "Situação do Pedido", "Valor do Pedido",
        "Valor da Condenação", "Observação do Pedido", "Adverso Principal",
        "Tipo do Adverso", "CPF/CNPJ do Adverso", "Escritório de Origem", "Escritório Responsável",
        "Centro de Custo", "Negociação do Contrato de Honorários"
    }

    campos_com_lookup_simples = {
        "Campo personalizado Citação",
        "Cliente Principal",
        "Fase",
        "Natureza",
        "Posição Cliente Principal",
        "Posição do Responsável Principal",
        "Procedimento",
        "Responsável Principal",
        "Resultado",
        "Risco",
        "Status do Processo",
        "Tipo de Ação",
        "Tipo de Contingência",
        "Tipo de Resultado",
        "Órgão"
    }

    campos_data_mascara_rigida = {
        "Data da Sentença",
        "Data da Baixa",
        "Data do Encerramento",
        "Data da Distribuição",
    }

    for coluna, valor in dados_linha.items():
        if coluna in campos_ja_tratados:
            continue
        if pd.isna(valor) or str(valor).strip() == "":
            continue

        campo = next((c for c in mapeamento if c.get("descricao") == coluna), None)
        if not campo or not campo.get("id"):
            campos_falha.append((coluna, "Elemento não mapeado"))
            continue

        try:
            campo_id = campo["id"]

            if coluna == "Campo personalizado Citação":
                seletor = 'input[id^="Citacao_ProcessoEntitySchema_"][id$="_Value"]'
                try:
                    page.wait_for_function(
                        """(selector) => {
                            const el = document.querySelector(selector);
                            if (!el) return false;
                            const wrapper = el.closest('.lookup');
                            return !el.readOnly && wrapper && !wrapper.classList.contains('disabled');
                        }""",
                        arg=seletor,
                        timeout=20000
                    )
                except:
                    campos_falha.append((coluna, "Timeout ao desbloquear campo"))
                    continue
                el = page.query_selector(seletor)

            elif coluna in {"Data da Terceirizacao", "Data da Terceirização"}:
                el = page.query_selector('input[id^="DataDeTerceirizacaoRecebimento_ProcessoEntitySchema_"]')

            elif coluna == "NPJ":
                el = page.query_selector('input[id^="NumeroDoCliente_ProcessoEntitySchema_"]')

            elif coluna == "Operações Vinculadas":
                el = page.query_selector('[id^="OperacoesVinculadas_ProcessoEntitySchema"]')

            else:
                el = page.query_selector(f'[id="{campo_id}"]') or page.query_selector(f'[id*="{campo_id}"]')

            if not el:
                campos_falha.append((coluna, "Elemento não encontrado"))
                continue

            if not el.is_enabled():
                campos_falha.append((coluna, "Campo desabilitado"))
                continue

            tag = el.evaluate("e => e.tagName.toLowerCase()")
            valor_formatado = valor.strftime("%d/%m/%Y") if isinstance(valor, pd.Timestamp) else str(valor)

            if coluna in campos_com_lookup_simples:
                sucesso = preencher_lookup_com_validacao(page, el, valor_formatado)
                if sucesso:
                    campos_sucesso.append(coluna)
                else:
                    campos_falha.append((coluna, "Falha ao validar lookup"))
                continue

            if tag == "input":
                if coluna == "Data do Resultado":
                    el.scroll_into_view_if_needed()
                    el.click()
                    el.press("Control+A")
                    el.press("Backspace")
                    el.type(valor_formatado, delay=100)
                    el.press("Tab")
                    page.wait_for_timeout(200)
                    print(f"  🗓️ Data do Resultado preenchida com: {valor_formatado}")

                elif coluna in campos_data_mascara_rigida:
                    el.click()
                    el.press("Home")
                    for _ in range(12):
                        el.press("Backspace")
                    el.type(valor_formatado, delay=100)
                    el.press("Tab")
                    page.wait_for_timeout(200)
                    print(f"  🗓️ {coluna} preenchida com: {valor_formatado}")

                elif "Data" in coluna:
                    el.click()
                    el.press("Control+A")
                    el.press("Backspace")
                    el.type(valor_formatado, delay=100)
                    el.press("Tab")
                    page.wait_for_timeout(200)
                    print(f"  🗓️ {coluna} preenchida com: {valor_formatado}")

                else:
                    el.click()
                    el.press("Control+A")
                    el.press("Backspace")
                    el.type(valor_formatado, delay=100)
                    el.press("Tab")
                campos_sucesso.append(coluna)

            elif tag == "select":
                el.select_option(label=valor_formatado)
                campos_sucesso.append(coluna)

            elif tag == "textarea":
                el.fill(valor_formatado)
                campos_sucesso.append(coluna)

            else:
                campos_falha.append((coluna, f"Tag não tratada: {tag}"))

        except Exception as e:
            campos_falha.append((coluna, str(e)))

    if campos_sucesso:
        print(f"  \033[92m✔ Campos gerais preenchidos com sucesso ({len(campos_sucesso)})\033[0m")
    if campos_falha:
        print(f"  \033[93m⚠️ {len(campos_falha)} campo(s) com falha no preenchimento:\033[0m")
        for campo, erro in campos_falha:
            print(f"    └─ ⚠️ {campo}: {erro}")

    return {
        "etapa": etapa,
        "duracao": round(time.time() - inicio, 2),
        "status": "Sucesso" if campos_sucesso else "Falha",
        "mensagem": "; ".join([f"{c}: {e}" for c, e in campos_falha]) if campos_falha else ""
    }

def preencher_adverso_principal(page, dados_linha):
    """PREENCHIMENTO DO ADVERSO PRINCIPAL"""
    import time
    inicio = time.time()
    etapa = "Adverso"

    try:
        nome_adverso = str(dados_linha.get("Adverso Principal", "")).strip()
        if not nome_adverso:
            return {
                "etapa": etapa,
                "duracao": round(time.time() - inicio, 2),
                "status": "Falha",
                "mensagem": "Campo 'Adverso Principal' vazio"
            }

        print(f" ├─ Preenchendo campo 'Contrário principal' com valor: {nome_adverso}")

        input_adverso = page.query_selector('input[id="Contrario_EnvolvidoText"]')
        if not input_adverso or not input_adverso.is_enabled():
            msg = "Campo 'Contrário principal' não localizado ou desabilitado."
            print(f"  ⚠️ {msg}")
            return {
                "etapa": etapa,
                "duracao": round(time.time() - inicio, 2),
                "status": "Falha",
                "mensagem": msg
            }

        input_adverso.click()
        input_adverso.fill(nome_adverso)
        page.wait_for_timeout(500)
        input_adverso.press("Enter")
        page.wait_for_timeout(600)
        input_adverso.press("ArrowDown")
        page.wait_for_timeout(800)

        sem_resultado = page.query_selector("div.empty-box")
        if sem_resultado:
            print("  ⚠️ Valor não existe no lookup. Abrindo modal de cadastro...")

            botao_plus = page.query_selector('#Contrario_lookup_envolvido .lookup-button.lookup-new')
            if botao_plus:
                botao_plus.click()
                page.wait_for_selector('form#contatoForm', timeout=10000)

                tipo = str(dados_linha.get("Tipo do Adverso", "Pessoa física")).strip()
                doc = str(dados_linha.get("CPF/CNPJ do Adverso", "")).strip()

                sucesso_modal = preencher_modal_adverso(page, nome_adverso, tipo, doc)
                if sucesso_modal:
                    return {
                        "etapa": etapa,
                        "duracao": round(time.time() - inicio, 2),
                        "status": "Sucesso"
                    }
                else:
                    return {
                        "etapa": etapa,
                        "duracao": round(time.time() - inicio, 2),
                        "status": "Falha",
                        "mensagem": "Falha ao cadastrar adverso no modal"
                    }
            else:
                msg = "Botão de criação (+) não encontrado."
                print(f"  ❌ {msg}")
                return {
                    "etapa": etapa,
                    "duracao": round(time.time() - inicio, 2),
                    "status": "Falha",
                    "mensagem": msg
                }
        else:
            input_adverso.press("Enter")
            page.wait_for_timeout(300)
            input_adverso.press("Tab")
            print("  ✅ Valor encontrado no lookup. Selecionado com sucesso.")
            return {
                "etapa": etapa,
                "duracao": round(time.time() - inicio, 2),
                "status": "Sucesso"
            }

    except Exception as e:
        return {
            "etapa": etapa,
            "duracao": round(time.time() - inicio, 2),
            "status": "Falha",
            "mensagem": f"Erro ao preencher 'Contrário principal': {e}"
        }

def preencher_modal_adverso(page, nome_adverso, tipo="Pessoa física", cpf_cnpj=""):

    """PREENCHIMENTO DO MODAL DE CADASTRO DO ADVERSO"""
    import time
    inicio = time.time()
    etapa = "Cadastro de Adverso"

    try:
        page.wait_for_selector('form#contatoForm', timeout=10000)

        page.fill('input#Nome', nome_adverso)
        print(f"  🖊️ Nome preenchido: {nome_adverso}")

        page.select_option('select#Tipo', label=tipo)
        print(f"  🖊️ Tipo selecionado: {tipo}")
        page.wait_for_timeout(500)

        if cpf_cnpj:
            if tipo == "Pessoa física":
                cpf_input = page.query_selector('input#CPF')
                if cpf_input and cpf_input.is_enabled():
                    cpf_input.fill(cpf_cnpj)
                    print(f"  🖊️ CPF preenchido: {cpf_cnpj}")
            elif tipo == "Pessoa jurídica":
                cnpj_input = page.query_selector('input#CNPJ')
                if cnpj_input and cnpj_input.is_visible():
                    cnpj_input.fill(cpf_cnpj)
                    print(f"  🖊️ CNPJ preenchido: {cpf_cnpj}")
        else:
            print("  ⚠️ Nenhum CPF/CNPJ informado — campo deixado em branco.")

        page.fill('input#Justificativa', 'rpa')

        page.click('input#btnSaveAddContact')
        print("  💾 Salvando cadastro do adverso...")
        page.wait_for_timeout(1500)

        return True  # ⬅️ Mantido por compatibilidade com função `preencher_adverso_principal`

    except Exception as e:
        print(f"[❌] Erro ao preencher modal do adverso: {e}")
        return False

def preencher_lookup_com_validacao(page, el_input, valor):

    try:
        if not el_input or not el_input.is_visible():
            return False

        tag = el_input.evaluate("el => el.tagName.toLowerCase()")
        if tag == "select":
            print("  ⚠️ Campo é um <select>, não requer validação de lookup.")
            return True

        is_enabled = el_input.is_enabled()
        handle = el_input.evaluate_handle("el => el")
        if not handle:
            print("  ⚠️ ElementHandle não disponível.")
            return False

        is_readonly = page.evaluate("el => el.readOnly", handle)
        if not is_enabled or is_readonly:
            print(f"  ⚠️ Campo desabilitado ou readonly.")
            return False

        el_input.scroll_into_view_if_needed()
        el_input.click()
        el_input.fill(str(valor))
        page.wait_for_timeout(600)

        el_input.press("Enter")
        page.wait_for_timeout(1000)
        el_input.press("ArrowDown")
        page.wait_for_timeout(1000)
        el_input.press("Enter")
        page.wait_for_timeout(1000)
        el_input.press("Tab")
        page.wait_for_timeout(1000)

        return True
    except Exception as e:
        print(f"[⚠️] Erro ao validar lookup: {e}")
        return False
    
def preencher_escritorio_responsavel(page, valor_lookup, dicionario):
    print(f" ├─ Preenchendo Escritório Responsável: {valor_lookup}")
    id_valor = dicionario.get(valor_lookup)
    if not id_valor:
        print(f"  ⚠️ Escritório não encontrado no dicionário: '{valor_lookup}'")
        return False

    try:
        # 1. Clica na lupa para ativar o lookup
        botao_lupa = page.locator("#LookupTreeEscritorioResponsavel .lookup-filter")
        botao_lupa.scroll_into_view_if_needed()
        botao_lupa.click()
        page.wait_for_timeout(1000)

        # 2. Clica no primeiro item da lista aberta
        primeiro_item = page.locator("table tbody td").first
        primeiro_item.click()
        page.wait_for_timeout(500)

        # 3. Seta o ID correto no campo oculto
        campo_hidden = page.locator("input#EscritorioResponsavelId")
        campo_hidden.evaluate(f"""
            (el) => {{
                el.value = '{id_valor}';
                el.dispatchEvent(new Event('change', {{ bubbles: true }}));
            }}
        """)

        print(f"  ✅ Escritório Responsável setado com ID {id_valor}")
        return True

    except Exception as e:
        print(f"  ❌ Erro ao preencher Escritório Responsável: {e}")
        return False

def preencher_negociacao_honorario(page, valor_lookup, dicionario):
    print(f" ├─ Preenchendo Negociação do Contrato de Honorário: {valor_lookup}")
    id_valor = dicionario.get(valor_lookup)
    if not id_valor:
        print(f"  ⚠️ Negociação não encontrada no dicionário: '{valor_lookup}'")
        return False

    try:
        botao_lupa = page.locator("#lookupgrid_negociacao .lookup-filter")
        botao_lupa.scroll_into_view_if_needed()
        botao_lupa.click()
        page.wait_for_timeout(1000)

        primeiro_item = page.locator("table tbody td").first
        primeiro_item.click()
        page.wait_for_timeout(500)

        campo_hidden = page.locator("input#NegociacaoContratoHonorarioId")
        campo_hidden.evaluate(f"""
            (el) => {{
                el.value = '{id_valor}';
                el.dispatchEvent(new Event('change', {{ bubbles: true }}));
            }}
        """)

        print(f"  ✅ Negociação setada com ID {id_valor}")
        return True

    except Exception as e:
        print(f"  ❌ Erro ao preencher Negociação: {e}")
        return False

def preencher_centro_custo(page, valor_lookup, dicionario):
    print(f" ├─ Preenchendo Centro de Custo: {valor_lookup}")
    id_valor = dicionario.get(valor_lookup)
    if not id_valor:
        print(f"  ⚠️ Centro de Custo não encontrado no dicionário: '{valor_lookup}'")
        return False

    try:
        botao_lupa = page.locator("div[id*='lookupTreeAreaRateio'] .lookup-filter")
        botao_lupa.scroll_into_view_if_needed()
        botao_lupa.click()
        page.wait_for_timeout(1000)

        primeiro_item = page.locator("table tbody td").first
        primeiro_item.click()
        page.wait_for_timeout(500)

        campo_hidden = page.locator("input[id*='AreaId']")
        campo_hidden.evaluate(f"""
            (el) => {{
                el.value = '{id_valor}';
                el.dispatchEvent(new Event('change', {{ bubbles: true }}));
            }}
        """)

        print(f"  ✅ Centro de Custo setado com ID {id_valor}")
        return True

    except Exception as e:
        print(f"  ❌ Erro ao preencher Centro de Custo: {e}")
        return False

def main():
    start_time = time.time()
    print("🟢 INÍCIO DO SCRIPT RPA", flush=True)

    usuario = os.environ.get("LEGALONE_USUARIO")
    senha = os.environ.get("LEGALONE_SENHA")
    if not usuario or not senha:
        print("❌ Variáveis de ambiente não definidas.")
        return

    BASE_DIR = Path(__file__).resolve().parent
    DATA_FILE = BASE_DIR / "data" / "entrada.xlsx"
    JSON_FILE = BASE_DIR / "data" / "campos_mapeados.json"

    try:
        df = pd.read_excel(DATA_FILE, sheet_name="Dados")
        with open(JSON_FILE, encoding="utf-8") as f:
            todos_campos = json.load(f)
        mapeamento = todos_campos
    except Exception as e:
        print(f"[❌] Erro ao carregar dados: {e}")
        return

    logs_processos = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        page = browser.new_page()

        # Login
        login_info = login_legalone(page, usuario, senha)
        print(f"[{login_info['etapa']}] {login_info['status']} ({login_info['duracao']}s)")
        if login_info["status"] == "Falha":
            print(f"❌ Erro no login: {login_info.get('mensagem', '')}")
            return

        for i, row in df.iterrows():
            dados_linha = row.dropna().to_dict()
            numero = str(dados_linha.get("Processo", "")).strip()
            if not numero:
                continue

            print(f"\n▶️ Atualizando processo: {numero}")
            log = iniciar_log_processo(numero)

            # Acesso
            acesso = acessar_processo_para_edicao(page, numero)
            print(f"[{acesso['etapa']}] {acesso['status']} ({acesso['duracao']}s)")
            adicionar_evento(log, acesso["etapa"], acesso["status"].lower(), "Abertura", acesso.get("mensagem", ""), acesso["duracao"])
            if acesso["status"] == "Falha":
                logs_processos.append(log)
                continue

            # Painéis
            expandir = expandir_paineis(page)
            adicionar_evento(log, expandir["etapa"], expandir["status"].lower(), "Expansão", expandir.get("mensagem", ""), expandir["duracao"])

            # Escritório Responsável
            valor_escritorio = dados_linha.get("Escritório Responsável")
            if valor_escritorio:
                inicio_escritorio = time.time()
                sucesso = preencher_escritorio_responsavel(page, valor_escritorio, MAPA_ESCRITORIO_RESPONSAVEL)
                duracao = round(time.time() - inicio_escritorio, 2)
                status = "sucesso" if sucesso else "erro"
                adicionar_evento(log, "Escritório Responsável", status, "Preenchimento", "", duracao)

            # Negociação do Contrato de Honorários
            valor_negociacao = dados_linha.get("Negociação do Contrato de Honorários")
            if valor_negociacao:
                inicio_negociacao = time.time()
                sucesso = preencher_negociacao_honorario(page, valor_negociacao, MAPA_NEGOCIACAO_HONORARIO)
                duracao = round(time.time() - inicio_negociacao, 2)
                status = "sucesso" if sucesso else "erro"
                adicionar_evento(log, "Negociação do Contrato de Honorários", status, "Preenchimento", "", duracao)

            # Centro de Custo
            valor_centro_custo = dados_linha.get("Centro de Custo")
            if valor_centro_custo:
                inicio_cc = time.time()
                sucesso = preencher_centro_custo(page, valor_centro_custo, MAPA_CENTRO_CUSTO)
                duracao = round(time.time() - inicio_cc, 2)
                status = "sucesso" if sucesso else "erro"
                adicionar_evento(log, "Centro de Custo", status, "Preenchimento", "", duracao)

            # Cascata
            cascata = preencher_lookups_em_cascata(page, dados_linha, mapeamento)
            adicionar_evento(log, cascata["etapa"], cascata["status"].lower(), "Lookups", cascata.get("mensagem", ""), cascata["duracao"])

            # Envolvidos
            envolvidos = preencher_outros_envolvidos(page, dados_linha)
            adicionar_evento(log, envolvidos["etapa"], envolvidos["status"].lower(), "Preenchimento", envolvidos.get("mensagem", ""), envolvidos["duracao"])

            # Objetos
            objetos = preencher_objetos(page, dados_linha)
            adicionar_evento(log, objetos["etapa"], objetos["status"].lower(), "Preenchimento", objetos.get("messagem", ""), objetos["duracao"])

            # Pedidos
            pedidos = preencher_pedidos(page, dados_linha, mapeamento)
            adicionar_evento(log, pedidos["etapa"], pedidos["status"].lower(), "Preenchimento", pedidos.get("mensagem", ""), pedidos["duracao"])

            # Adverso principal (condicional)
            if "Adverso Principal" in dados_linha and str(dados_linha["Adverso Principal"]).strip():
                adverso = preencher_adverso_principal(page, dados_linha)
                adicionar_evento(log, adverso["etapa"], adverso["status"].lower(), "Principal", adverso.get("mensagem", ""), adverso["duracao"])

            # Campos gerais
            campos = atualizar_campos_usando_mapeamento(page, dados_linha, mapeamento)
            adicionar_evento(log, campos["etapa"], campos["status"].lower(), "Atualização", campos.get("messagem", ""), campos["duracao"])

            input("🔎 Valide os campos gerais e pressione Enter para continuar...")

            # SALVAR ALTERAÇÕES COM VERIFICAÇÃO DE FALHA
            try:
                inicio_salvar = time.time()
                url_antes = page.url
                botao_salvar = page.query_selector('button[name="ButtonSave"]')

                if botao_salvar and botao_salvar.is_enabled():
                    botao_salvar.scroll_into_view_if_needed()
                    botao_salvar.click()
                    page.wait_for_timeout(2000)

                    url_depois = page.url
                    duracao = round(time.time() - inicio_salvar, 2)

                    if url_antes == url_depois:
                        mensagem = "Falha ao salvar — validação obrigatória não atendida"
                        adicionar_evento(log, "Salvar", "erro", "Confirmação", mensagem, duracao)
                        print(f"[❌] {mensagem} no processo {numero}")
                    else:
                        adicionar_evento(log, "Salvar", "sucesso", "Confirmação", "", duracao)
                        print(f"💾 Processo {numero} salvo com sucesso ({duracao}s)")
                else:
                    adicionar_evento(log, "Salvar", "erro", "Confirmação", "Botão desabilitado ou não encontrado", 0)
                    print(f"[⚠️] Botão 'Salvar' não disponível para o processo {numero}")
            except Exception as e:
                adicionar_evento(log, "Salvar", "erro", "Confirmação", str(e), 0)
                print(f"[ERRO] Falha ao tentar salvar processo {numero}: {e}")

            print(f"[✔️] Preenchimento concluído para {numero}.")
            logs_processos.append(log)

        input("\n🧪 Pressione Enter para encerrar o navegador...")

    salvar_log_execucao(logs_processos)
    tempo_total = time.time() - start_time
    print(f"\n⏱️ Tempo total de execução: {tempo_total:.2f} segundos.")

if __name__ == "__main__":
    main()