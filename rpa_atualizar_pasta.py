from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
from pathlib import Path
from datetime import datetime
import pandas as pd
import json
import os
import time
import sys
import argparse
import re

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

        # Garante que o robô sempre clique no ícone da PRIMEIRA linha da tabela de resultados
        print("   - Focando no primeiro resultado da busca...")
        page.locator("table.webgrid tbody tr").first.locator("span.grid-overflow-icon").click()
        
        # Clica na primeira ação de "Editar" que aparecer
        page.locator("a.grid-edit-action-row").first.click()
        
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
    """
    PREENCHIMENTO DOS CAMPOS EM CASCATA COM SELEÇÃO EXATA DE LOOKUP PELA SIGLA (UF).
    Remove chamadas a .is_focused() que causavam erro.
    """
    inicio = time.time()
    etapa = "Cascata"
    mensagens_erro = []
    campos_preenchidos = 0

    print(" - Preenchendo campos em cascata (UF -> Cidade -> ... -> Tipo da Vara)...")

    ordem = [
        "Estado (UF)", "Cidade", "Justiça (CNJ)", "Instância (CNJ)",
        "Classe (CNJ)", "Assunto (CNJ)", "Comarca/Foro", "Número da Vara",
        "Tipo da Vara"
    ]

    for atual in ordem:
        try:
            valor = dados_linha.get(atual)
            if pd.isna(valor) or not str(valor).strip():
                print(f"    [AVISO] Campo '{atual}' sem valor na planilha, pulando.")
                continue

            valor_str = str(valor).strip()

            campo = next((c for c in mapeamento if c.get("descricao") == atual), None)
            if not campo or not campo.get("id"):
                print(f"    [AVISO] Campo '{atual}' não mapeado no JSON.")
                continue

            # --- Tratamento para "Número da Vara" ---
            if atual == "Número da Vara":
                seletor_vara = f'input[id="{campo["id"]}"], input[id*="{campo["id"]}"]'
                num_vara_locator = page.locator(seletor_vara).first
                try:
                    num_vara_locator.fill(valor_str, timeout=10000)
                    num_vara_locator.press("Tab")
                    page.wait_for_timeout(300)
                    print(f"    [OK] Campo '{atual}' preenchido: {valor_str}")
                    campos_preenchidos += 1
                except PlaywrightTimeoutError:
                     print(f"    [ERRO] Timeout ao tentar preencher '{atual}'.")
                     mensagens_erro.append(f"{atual}: Timeout no fill")
                except Exception as e_vara:
                    print(f"    [ERRO] Erro ao preencher '{atual}': {e_vara}")
                    mensagens_erro.append(f"{atual}: {e_vara}")
                continue

            # --- Tratamento para campos Lookup ---
            locator = None
            seletor_str = ""

            if atual == "Assunto (CNJ)":
                identificador = campo["id"].split("__")[-1]
                seletor_str = f'input[id^="Assuntos_"][id$="__{identificador}"]:not([type="hidden"])'
            else:
                seletor_str = f'input.search.ac_input[id="{campo["id"]}"], input.search.ac_input[id*="{campo["id"]}"]'

            locator = page.locator(seletor_str).first

            # --- Lógica de Preenchimento e Seleção Exata ---
            print(f"    [AÇÃO] Preenchendo '{atual}' com valor: '{valor_str}'")
            try:
                # Clica para focar e espera habilitar (com timeout na ação)
                locator.click(timeout=15000)
                if not locator.is_enabled(): # Verifica APÓS espera implícita do click
                     raise Exception("Campo não está habilitado após clique/espera.")

                print(f"    [INFO] Campo '{atual}' clicado e habilitado.")
                locator.fill(valor_str)
                page.wait_for_timeout(800)
                locator.press("Enter")
                page.wait_for_timeout(1500) # Espera lista carregar

                item_exato_clicado = False
                try:
                    # Tenta selecionar o item exato na lista
                    suggestion_table_selector = 'div.lookup-dropdown:visible tbody tr'
                    page.locator(suggestion_table_selector).first.wait_for(state='visible', timeout=8000)

                    if atual == "Estado (UF)":
                         exact_match_selector = f'{suggestion_table_selector}:has(td[data-val-field="UFSigla"]:text-is("{valor_str}"))'
                         data_val_field_log = "UFSigla"
                    else:
                         data_val_field_nome = campo.get("id").replace("Text", "")
                         exact_match_selector = f'{suggestion_table_selector}:has(td[data-val-field="{data_val_field_nome}Text"]:text-is("{valor_str}"))'
                         data_val_field_log = f"{data_val_field_nome}Text"
                         if page.locator(exact_match_selector).count() == 0:
                              exact_match_selector = f'{suggestion_table_selector}:has(td:text-is("{valor_str}"))'
                              data_val_field_log = "Texto Genérico da Célula"

                    print(f"      - Procurando linha com '{data_val_field_log}' exatamente igual a '{valor_str}'...")
                    exact_match_locator = page.locator(exact_match_selector)

                    if exact_match_locator.count() > 0:
                        print(f"      - Encontrada correspondência exata para '{valor_str}'. Clicando na linha...")
                        exact_match_locator.first.click()
                        item_exato_clicado = True
                        page.wait_for_timeout(500)
                    else:
                        print(f"      - [AVISO] Correspondência exata para '{valor_str}' (via {data_val_field_log}) não encontrada na lista visível.")

                except PlaywrightTimeoutError:
                     print(f"      - [AVISO] Lista de sugestões não apareceu a tempo para '{atual}'.")
                except Exception as e_select:
                    print(f"      - [AVISO] Problema ao processar lista de sugestões (tabela): {e_select}")

                # Fallback se não conseguiu clicar no item exato
                if not item_exato_clicado:
                    print("      - Usando fallback: ArrowDown + Enter.")
                    try:
                        # <<< CORREÇÃO: Removido is_focused() >>>
                        locator.press("ArrowDown")
                        page.wait_for_timeout(400)
                        locator.press("Enter")
                        page.wait_for_timeout(400)
                    except Exception as e_fallback:
                         # Captura erro se não conseguir pressionar (ex: elemento não focado)
                         print(f"      - [ERRO] Erro durante fallback ArrowDown+Enter: {e_fallback}")
                         # Mesmo com erro no fallback, tenta sair com Tab

                # Sempre tenta pressionar Tab para sair do campo
                try:
                    # <<< CORREÇÃO: Removido is_focused() >>>
                    locator.press("Tab")
                    page.wait_for_timeout(500)
                except Exception as e_tab:
                     print(f"      - [AVISO] Erro ao tentar sair do campo com Tab: {e_tab}. Tentando clicar fora.")
                     # Fallback clicando no body se Tab falhar
                     try:
                          page.locator('body').click()
                          page.wait_for_timeout(300)
                     except Exception as e_click_body:
                          print(f"      - [ERRO] Falha ao clicar fora do campo: {e_click_body}")


                print(f"    [OK] Campo '{atual}' preenchido.")
                campos_preenchidos += 1

            # Captura erro durante o preenchimento (click, fill, seleção)
            except PlaywrightTimeoutError as e_timeout:
                 print(f"    [ERRO] Timeout ao interagir com '{atual}': {e_timeout}")
                 mensagens_erro.append(f"{atual}: Timeout na interação ({e_timeout})")
            except Exception as e_fill:
                 print(f"    [ERRO] Erro durante o preenchimento/seleção de '{atual}': {e_fill}")
                 mensagens_erro.append(f"{atual}: {e_fill}")


        # Captura erro geral para o item 'atual' do loop
        except Exception as e:
            print(f"    [ERRO GERAL] Erro inesperado ao processar '{atual}': {e}")
            mensagens_erro.append(f"{atual}: {e}")

    # --- Retorno da Função ---
    # (Lógica de retorno mantida igual à anterior)
    if campos_preenchidos > 0:
        status_final = "Sucesso"
        if mensagens_erro:
             print(f"  [AVISO] {len(mensagens_erro)} erro(s) ocorreram durante preenchimento da cascata.")
        return {
            "etapa": etapa, "duracao": round(time.time() - inicio, 2),
            "status": status_final, "mensagem": "; ".join(mensagens_erro) if mensagens_erro else ""
        }
    else:
        return {
            "etapa": etapa, "duracao": round(time.time() - inicio, 2),
            "status": "Falha" if mensagens_erro else "Sucesso",
            "mensagem": "; ".join(mensagens_erro) if mensagens_erro else "Nenhum campo da cascata precisou ser preenchido."
        }

def preencher_outros_envolvidos(page, dados_linha):
    """PREENCHIMENTO DOS ENVOLVIDOS"""
    import time
    inicio = time.time()
    etapa = "Envolvidos"
    erros = []
    sucessos = 0

    print(" - Preenchendo Envolvidos...")

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
                print(f"[AVISO] {msg}")
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
                    print(f"  - Situação {i+1}: {situacoes[i]}")
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
                            print(f"  - Posição {i+1}: {posicoes[i]}")
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
                        print(f"  [AVISO] Envolvido {i+1} não encontrado. Abrindo modal...")
                        botao_plus = bloco_lookup.locator('.lookup-button.lookup-new')
                        if botao_plus:
                            botao_plus.click()
                            page.wait_for_selector('form#contatoForm', timeout=10000)
                            sucesso_modal = preencher_modal_adverso(page, nome, tipo, doc)
                            if sucesso_modal:
                                print(f"  [SUCESSO] Envolvido {i+1} cadastrado com sucesso.")
                                sucessos += 1
                            else:
                                erros.append(f"Falha ao cadastrar Envolvido {i+1}")
                        else:
                            erros.append(f"Botão '+' do Envolvido {i+1} não encontrado")
                    else:
                        input_nome.first.press("Enter")
                        page.wait_for_timeout(300)
                        input_nome.first.press("Tab")
                        print(f"  - Nome {i+1}: {nome}")
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

    print("  - Preenchendo Objetos...")
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
                print(f"[AVISO] {msg}")
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
                        print(f"  - Nome do Objeto {i+1}: {nomes[i]}")
                        sucessos += 1
                    else:
                        msg = f"Validação falhou para Nome do Objeto {i+1}"
                        print(f"[AVISO] {msg}")
                        erros.append(msg)
                except Exception as e:
                    erros.append(f"Erro Nome {i+1}: {e}")

            if i < len(observacoes):
                try:
                    grupo.locator('textarea[id*="Observacoes"]').fill(observacoes[i])
                    print(f"  - Observações do Objeto {i+1}")
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
        print(f"[AVISO] Só {visiveis} blocos visíveis após {timeout_ms}ms (esperado: {quantidade_esperada})")
        return False

    print(" - Preenchendo Pedidos...")
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
    print(f"  - Total detectado: {num_pedidos} pedido(s)")

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
    print(f"  - {pedidos_li.count()} blocos de pedido visíveis")

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
                        print(f"  [INFO] Data do Julgamento preenchida com: {valor}")
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
            print(f"   - [SUCESSO] Pedido {i+1} preenchido com sucesso.")

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
    """PREENCHIMENTO DOS CAMPOS GERAIS (COM CORREÇÃO DE VÍRGULA DECIMAL)"""
    inicio = time.time()
    etapa = "Campos Gerais"
    campos_sucesso = []
    campos_falha = []

    print(" - Preenchendo campos gerais...")

    campos_ja_tratados = {
        "Processo", "Estado (UF)", "Cidade", "Justiça (CNJ)", "Instância (CNJ)",
        "Classe (CNJ)", "Assunto (CNJ)", "Situação do Envolvido", "Posição Envolvido",
        "Envolvido", "Nome do Objeto", "Observações do objeto", "Tipo de Pedido",
        "Contingência", "Data do Pedido", "Data do Julgamento", "Probabilidade de Êxito",
        "Situação do Pedido", "Valor do Pedido", "Valor da Condenação",
        "Observação do Pedido", "Adverso Principal", "Tipo do Adverso",
        "CPF/CNPJ do Adverso", "Escritório de Origem", "Escritório Responsável",
        "Centro de Custo", "Negociação do Contrato de Honorários"
    }

    campos_com_lookup_simples = {
        "Campo personalizado Citação", "Cliente Principal", "Fase", "Natureza",
        "Posição Cliente Principal", "Posição do Responsável Principal", "Procedimento",
        "Responsável Principal", "Resultado", "Risco", "Status do Processo",
        "Tipo de Ação", "Tipo de Contingência", "Tipo de Resultado", "Órgão"
    }

    campos_data_mascara_rigida = {
        "Data da Sentença", "Data da Baixa", "Data do Encerramento",
        "Data da Distribuição", "Data da Terceirização", "Data da Terceirizacao"
    }

    # Identifica colunas que tipicamente contêm valores numéricos/monetários
    colunas_de_valor = {
        "Valor da Causa",
        "Valor da Condenação/Acordo",
        "Valor do Pedido", # Incluindo caso apareçam aqui
        "Valor da Condenação" # Incluindo caso apareçam aqui
        # Adicione outras colunas de valor se existirem
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
            el = None # Inicializa

            # --- Lógica original para encontrar o elemento (el) ---
            # (Mantendo a lógica original de busca com query_selector por enquanto,
            # pois o problema principal é a formatação do valor)
            if coluna == "Campo personalizado Citação":
                seletor = 'input[id^="Citacao_ProcessoEntitySchema_"][id$="_Value"]'
                try:
                    page.wait_for_function(
                        """(selector) => { /* ... sua função JS ... */ }""",
                        arg=seletor, timeout=20000
                    )
                    el = page.query_selector(seletor)
                except Exception as e_wait:
                    campos_falha.append((coluna, f"Timeout/Erro ao esperar desbloqueio JS: {e_wait}"))
                    continue
            elif coluna in {"Data da Terceirizacao", "Data da Terceirização"}:
                el = page.query_selector('input[id^="DataDeTerceirizacaoRecebimento_ProcessoEntitySchema_"]')
            elif coluna == "NPJ":
                el = page.query_selector('input[id^="NumeroDoCliente_ProcessoEntitySchema_"]')
            elif coluna == "Operações Vinculadas":
                el = page.query_selector('[id^="OperacoesVinculadas_ProcessoEntitySchema"]')
            else:
                el = page.query_selector(f'[id="{campo_id}"]') or page.query_selector(f'[id*="{campo_id}"]:not([type="hidden"])')

            # --- Verificações do elemento ---
            if not el:
                campos_falha.append((coluna, "Elemento não encontrado"))
                continue
            if not el.is_enabled(): # is_enabled() funciona em ElementHandle
                campos_falha.append((coluna, "Campo desabilitado"))
                continue

            # --- Lógica de formatação e preenchimento ---
            tag = el.evaluate("e => e.tagName.toLowerCase()")

            # --- AJUSTE NA FORMATAÇÃO DO VALOR ---
            valor_formatado = ""
            if isinstance(valor, pd.Timestamp):
                valor_formatado = valor.strftime("%d/%m/%Y")
            # Verifica se a coluna é de valor E se o tipo é numérico
            elif coluna in colunas_de_valor and isinstance(valor, (int, float)):
                 # Formata para string, usando vírgula como decimal, SEM separador de milhar por ora
                 # Se o campo web espera separador de milhar, a formatação precisa ser mais complexa
                 valor_formatado = f"{valor:.2f}".replace('.', ',') # Garante 2 casas decimais e usa vírgula
                 # Se for inteiro, o .2f adiciona ',00', o que pode ser indesejado.
                 # Tratamento para remover ',00' de inteiros, se necessário:
                 if isinstance(valor, int):
                      valor_formatado = str(valor) # Inteiro não precisa de vírgula decimal
            else:
                # Para todos os outros tipos, apenas converte para string
                valor_formatado = str(valor)
            # --- FIM DO AJUSTE NA FORMATAÇÃO ---


            if coluna in campos_com_lookup_simples:
                sucesso = preencher_lookup_com_validacao(page, el, valor_formatado)
                if sucesso:
                    campos_sucesso.append(coluna)
                    print(f"    [OK] Lookup '{coluna}' preenchido.")
                else:
                    campos_falha.append((coluna, "Falha ao validar/preencher lookup"))
                continue

            if tag == "input":
                # Mantém a lógica original de limpar com Ctrl+A e Backspace
                try:
                    el.click()
                    el.press("Control+A")
                    el.press("Backspace")

                    # Lógica específica para datas com máscara
                    if coluna in campos_data_mascara_rigida:
                        el.press("Home") # Vai para o início após limpar
                        el.type(valor_formatado, delay=100)
                        print(f"    [OK] Data (máscara) '{coluna}' preenchida: {valor_formatado}")
                    # Lógica para outros inputs (incluindo os de valor agora formatados)
                    else:
                        el.type(valor_formatado, delay=80) # Digita o valor JÁ FORMATADO
                        if coluna in colunas_de_valor:
                             print(f"    [OK] Valor '{coluna}' preenchido: {valor_formatado}")
                        elif "Data" in coluna: # Para datas sem máscara rígida
                             print(f"    [OK] Data '{coluna}' preenchida: {valor_formatado}")
                        else:
                             print(f"    [OK] Campo '{coluna}' preenchido: {valor_formatado}")

                    el.press("Tab")
                    page.wait_for_timeout(250)
                    campos_sucesso.append(coluna)

                except Exception as e_input:
                    campos_falha.append((coluna, f"Erro ao preencher input: {e_input}"))

            elif tag == "select":
                try:
                    el.select_option(label=valor_formatado)
                    campos_sucesso.append(coluna)
                    print(f"    [OK] Select '{coluna}' selecionado: {valor_formatado}")
                    page.wait_for_timeout(250)
                except Exception as e_select:
                     try: # Fallback por valor
                         el.select_option(value=valor_formatado)
                         campos_sucesso.append(coluna)
                         print(f"    [OK] Select '{coluna}' selecionado (por valor): {valor_formatado}")
                         page.wait_for_timeout(250)
                     except Exception as e_select2:
                          campos_falha.append((coluna, f"Erro ao selecionar option (label/valor): {e_select} / {e_select2}"))


            elif tag == "textarea":
                try:
                    el.fill(valor_formatado) # Fill deve limpar antes
                    campos_sucesso.append(coluna)
                    print(f"    [OK] Textarea '{coluna}' preenchido.")
                    page.wait_for_timeout(250)
                except Exception as e_textarea:
                    campos_falha.append((coluna, f"Erro ao preencher textarea: {e_textarea}"))

            else:
                campos_falha.append((coluna, f"Tag não tratada: {tag}"))

        except Exception as e_geral:
            campos_falha.append((coluna, f"Erro geral processando coluna: {str(e_geral)}"))

    # --- Resumo final ---
    # (A lógica de resumo e retorno permanece a mesma)
    if campos_sucesso:
        print(f"  [SUCESSO] Campos gerais: {len(campos_sucesso)} preenchidos.")
    if campos_falha:
        print(f"  [AVISO] Campos gerais: {len(campos_falha)} com falha:")
        for campo, erro in campos_falha:
            erro_msg = (str(erro)[:150] + '...') if len(str(erro)) > 150 else str(erro)
            print(f"    - [FALHA] {campo}: {erro_msg}")

    return {
        "etapa": etapa,
        "duracao": round(time.time() - inicio, 2),
        "status": "Sucesso" if campos_sucesso else "Falha",
        "mensagem": "; ".join([f"{c}: {str(e)[:100]}" for c, e in campos_falha]) if campos_falha else ""
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

        print(f" - Preenchendo campo 'Contrário principal' com valor: {nome_adverso}")

        input_adverso = page.query_selector('input[id="Contrario_EnvolvidoText"]')
        if not input_adverso or not input_adverso.is_enabled():
            msg = "Campo 'Contrário principal' não localizado ou desabilitado."
            print(f"  [AVISO] {msg}")
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
            print("  [AVISO] Valor não existe no lookup. Abrindo modal de cadastro...")

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
                print(f"  [ERRO] {msg}")
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
            print("  [SUCESSO] Valor encontrado no lookup. Selecionado com sucesso.")
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
        print(f"  [INFO] Nome preenchido: {nome_adverso}")

        page.select_option('select#Tipo', label=tipo)
        print(f"  [INFO] Tipo selecionado: {tipo}")
        page.wait_for_timeout(500)

        if cpf_cnpj:
            if tipo == "Pessoa física":
                cpf_input = page.query_selector('input#CPF')
                if cpf_input and cpf_input.is_enabled():
                    cpf_input.fill(cpf_cnpj)
                    print(f"  [INFO] CPF preenchido: {cpf_cnpj}")
            elif tipo == "Pessoa jurídica":
                cnpj_input = page.query_selector('input#CNPJ')
                if cnpj_input and cnpj_input.is_visible():
                    cnpj_input.fill(cpf_cnpj)
                    print(f"  [INFO] CNPJ preenchido: {cpf_cnpj}")
        else:
            print("  [AVISO] Nenhum CPF/CNPJ informado — campo deixado em branco.")

        page.fill('input#Justificativa', 'rpa')

        page.click('input#btnSaveAddContact')
        print("  [INFO] Salvando cadastro do adverso...")
        page.wait_for_timeout(1500)

        return True

    except Exception as e:
        print(f"[ERRO] Erro ao preencher modal do adverso: {e}")
        return False

def preencher_lookup_com_validacao(page, el_input, valor):

    try:
        if not el_input or not el_input.is_visible():
            return False

        tag = el_input.evaluate("el => el.tagName.toLowerCase()")
        if tag == "select":
            print("  [AVISO] Campo é um <select>, não requer validação de lookup.")
            return True

        is_enabled = el_input.is_enabled()
        handle = el_input.evaluate_handle("el => el")
        if not handle:
            print("  [AVISO] ElementHandle não disponível.")
            return False

        is_readonly = page.evaluate("el => el.readOnly", handle)
        if not is_enabled or is_readonly:
            print(f"  [AVISO] Campo desabilitado ou readonly.")
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
        print(f"[AVISO] Erro ao validar lookup: {e}")
        return False
    
def preencher_escritorio_responsavel(page, valor_lookup, dicionario):
    print(f" - Preenchendo Escritório Responsável: {valor_lookup}")
    id_valor = dicionario.get(valor_lookup)
    if not id_valor:
        print(f"  [AVISO] Escritório não encontrado no dicionário: '{valor_lookup}'")
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

        print(f"  [SUCESSO] Escritório Responsável setado com ID {id_valor}")
        return True

    except Exception as e:
        print(f"  [ERRO] Erro ao preencher Escritório Responsável: {e}")
        return False

def preencher_negociacao_honorario(page, valor_lookup, dicionario):
    print(f" - Preenchendo Negociação do Contrato de Honorário: {valor_lookup}")
    id_valor = dicionario.get(valor_lookup)
    if not id_valor:
        print(f"  [AVISO] Negociação não encontrada no dicionário: '{valor_lookup}'")
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

        print(f"  [SUCESSO] Negociação setada com ID {id_valor}")
        return True

    except Exception as e:
        print(f"  [ERRO] Erro ao preencher Negociação: {e}")
        return False

def preencher_centro_custo(page, valor_lookup, dicionario):
    print(f" - Preenchendo Centro de Custo: {valor_lookup}")
    id_valor = dicionario.get(valor_lookup)
    if not id_valor:
        print(f"  [AVISO] Centro de Custo não encontrado no dicionário: '{valor_lookup}'")
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

        print(f"  [SUCESSO] Centro de Custo setado com ID {id_valor}")
        return True

    except Exception as e:
        print(f"  [ERRO] Erro ao preencher Centro de Custo: {e}")
        return False
   
def lidar_com_popup_de_confirmacao(page):
    """
    Verifica se popups de confirmação ('Atenção' ou 'Monitoramentos')
    estão visíveis e clica no botão de confirmação ('#popup_ok').
    Usa um timeout configurável. Retorna True se tratado, False caso contrário.
    """
    print(" - Verificando se existe popup de confirmação...")
    popup_timeout = 7000 # Aumentei um pouco o timeout para dar mais margem
    try:
        # O locator busca pelo ID do botão de confirmação, comum a ambos os popups
        sim_button_locator = page.locator('input#popup_ok')

        # Espera o botão ficar visível (seja qual for o popup)
        sim_button_locator.wait_for(state='visible', timeout=popup_timeout)

        # Se chegou aqui, o botão está visível. Loga a mensagem do popup se possível.
        try:
            # Tenta pegar a mensagem específica do popup para log
            popup_message_locator = page.locator('#popup_message')
            if popup_message_locator.is_visible(timeout=500): # Verifica rápido se a mensagem existe
                 popup_text = popup_message_locator.text_content(timeout=500)
                 print(f"   - Popup detectado com mensagem: '{popup_text[:100]}...'. Clicando em 'Sim/Salvar processo'.")
            else:
                 print("   - Popup (sem mensagem específica lida) detectado. Clicando em 'Sim/Salvar processo'.")
        except Exception as e_msg:
             print(f"   - Popup detectado (erro ao ler msg: {e_msg}). Clicando em 'Sim/Salvar processo'.")

        sim_button_locator.click()

        # Espera o container geral do popup desaparecer
        # Usar um seletor que abranja ambos os possíveis containers, se necessário,
        # mas 'div#popup_container' (se for o container externo) pode ser suficiente.
        # Se 'popup_content' estiver dentro de 'popup_container', esperar por 'popup_container' basta.
        page.locator('div#popup_container').wait_for(state='hidden', timeout=5000)
        print("   - Popup de confirmação tratado com sucesso.")
        return True # Indica que o popup foi tratado

    except Exception as e: # Captura TimeoutError do wait_for ou outros erros
        if "Timeout" in str(e):
             # Mensagem mais genérica, pois pode ser qualquer um dos popups
             print(f"   - Nenhum botão de confirmação ('input#popup_ok') encontrado em {popup_timeout/1000}s.")
        else:
             print(f"   - [AVISO] Ocorreu um erro inesperado ao tentar lidar com o popup: {e}")
        return False # Indica que o popup NÃO foi tratado

def gerar_relatorio_sumarizado(lista_logs: list, timestamp: str):
    """
    Cria um relatório sumarizado em CSV, priorizando o erro mais relevante
    e incluindo a Linha Excel.
    """
    print("[INFO] Gerando relatório sumarizado em CSV...")
    if not lista_logs:
        print("[AVISO] Lista de logs vazia. Nenhum relatório sumarizado será gerado.")
        return

    dados_relatorio = []
    for log in lista_logs:
        processo = log.get("Processo", "N/A")
        status = log.get("Status Geral", "Indefinido")
        duracao_total = sum(log.get("Duracao por etapa (s)", {}).values())
        
        # --- INÍCIO DA MODIFICAÇÃO DE LÓGICA DE ERRO ---
        etapa_falha = ""
        mensagem_erro = ""
        
        if status == "Falha" and log.get("Erros"):
            lista_erros = log.get("Erros", [])
            
            # 1. Procura por um erro específico da etapa "Salvar" (nossa regra de negócio)
            erro_salvar = next((err for err in lista_erros if err.get("etapa") == "Salvar"), None)
            
            if erro_salvar:
                # Prioridade 1: O erro de "Salvar"
                etapa_falha = erro_salvar.get("etapa", "Salvar")
                mensagem_erro = erro_salvar.get("mensagem", "Erro ao salvar")
            else:
                # Prioridade 2: O último erro que ocorreu no processo
                ultimo_erro = lista_erros[-1]
                etapa_falha = ultimo_erro.get("etapa", "N/A")
                mensagem_erro = ultimo_erro.get("mensagem", "Sem detalhes")
        
        elif status == "Falha":
            # Caso raro: Status é Falha (porque não salvou), mas a lista de Erros está vazia
            etapa_falha = "Salvar"
            mensagem_erro = "Processo não foi salvo (etapa 'Salvar' não registrou 'sucesso'), mas nenhum erro específico foi capturado."
            
        # --- FIM DA MODIFICAÇÃO ---

        dados_relatorio.append({
            "Linha_Excel": log.get("Linha Excel", "N/A"), # <-- INCLUSÃO DA LINHA EXCEL
            "Processo": processo,
            "Status": status,
            "Duracao_Total_s": round(duracao_total, 2),
            "Etapa_Falha": etapa_falha,
            "Mensagem_Erro": mensagem_erro
        })

    # Cria a pasta para os relatórios sumarizados
    pasta_relatorios = Path(__file__).resolve().parent / "data" / "relatorios_sumarizados"
    pasta_relatorios.mkdir(parents=True, exist_ok=True)
    caminho_csv = pasta_relatorios / f"relatorio_{timestamp}.csv"

    # Cria e salva o DataFrame
    df_relatorio = pd.DataFrame(dados_relatorio)
    
    # Reordenando colunas para a Linha Excel vir primeiro
    colunas_ordenadas = ["Linha_Excel", "Processo", "Status", "Duracao_Total_s", "Etapa_Falha", "Mensagem_Erro"]
    df_relatorio = df_relatorio[colunas_ordenadas]
    
    df_relatorio.to_csv(caminho_csv, index=False, sep=';', encoding='utf-8-sig')
    
    print(f"[RELATÓRIO] Relatório sumarizado salvo em: {caminho_csv}")

def gerar_arquivo_reprocessamento(lista_logs: list, df_original: pd.DataFrame, timestamp: str):
    """
    Filtra o DataFrame original e salva um novo arquivo Excel
    apenas com as linhas que falharam, prontas para reprocessar.
    
    Argumentos:
        lista_logs (list): A lista completa de logs de processo.
        df_original (pd.DataFrame): O DataFrame carregado do 'entrada.xlsx' original.
        timestamp (str): A data/hora formatada para nomear o arquivo.
    """
    print("[INFO] Gerando arquivo de reprocessamento...")
    
    # 1. Pega a "Linha Excel" de todos os logs que falharam
    linhas_falhas = [
        log.get("Linha Excel") 
        for log in lista_logs 
        if log.get("Status Geral") == "Falha" and log.get("Linha Excel")
    ]
    
    if not linhas_falhas:
        print("[INFO] Nenhum processo falhou. Arquivo de reprocessamento não gerado.")
        return

    # 2. Converte números de linha (base 1, ex: 10) para índices do DataFrame (base 0, ex: 8)
    #    Como seu log usa 'excel_row_num = i + 2', o índice pandas (i) é 'excel_row_num - 2'
    #    Filtramos índices que podem estar fora dos limites do df_original
    max_index = len(df_original) - 1
    indices_falhos = [
        idx for idx in (linha - 2 for linha in linhas_falhas) 
        if 0 <= idx <= max_index
    ]

    if not indices_falhos:
        print("[AVISO] Falhas encontradas no log, mas índices de linha não correspondem ao DataFrame original.")
        return

    # 3. Filtra o DataFrame original para pegar apenas as linhas que falharam
    df_falhas = df_original.iloc[indices_falhos].copy()
    
    # 4. Define o caminho da pasta
    pasta_reprocessamento = Path(__file__).resolve().parent / "data" / "reprocessamento"
    pasta_reprocessamento.mkdir(parents=True, exist_ok=True)
    caminho_excel = pasta_reprocessamento / f"falhas_{timestamp}.xlsx"
    
    # 5. Salva o novo arquivo Excel
    #    É crucial salvar na aba "Dados", pois é a aba que o 'main' lê.
    try:
        with pd.ExcelWriter(caminho_excel, engine='openpyxl') as writer:
            df_falhas.to_excel(writer, sheet_name='Dados', index=False)
            
        print(f"[REPROCESSAR] Arquivo com {len(df_falhas)} falhas salvo em: {caminho_excel}")
        print(f"[AÇÃO] Para reprocessar: Renomeie '{caminho_excel.name}' para 'entrada.xlsx' e mova para a pasta 'data'.")

    except Exception as e:
        print(f"[ERRO] Falha ao gerar arquivo de reprocessamento: {e}")

def salvar_log_execucao(lista_logs: list, df_original: pd.DataFrame):
    """
    Salva o log detalhado em JSON, chama o relatório sumarizado
    e gera o arquivo de reprocessamento.
    
    Argumentos:
        lista_logs (list): A lista de logs da execução.
        df_original (pd.DataFrame): O DataFrame completo do 'entrada.xlsx'.
    """
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    pasta_logs = Path(__file__).resolve().parent / "data" / "logs_execucao"
    pasta_logs.mkdir(parents=True, exist_ok=True)
    caminho_json = pasta_logs / f"log_{timestamp}.json"

    # Atualiza status geral com base na regra de negócio final
    # (Sua lógica original está perfeita)
    for log in lista_logs:
        sucesso_ao_salvar = any(
            evento.get("etapa") == "Salvar" 
            for evento in log.get("Sucessos", [])
        )

        if sucesso_ao_salvar:
            log["Status Geral"] = "Sucesso"
        else:
            log["Status Geral"] = "Falha"

    # Salva o log detalhado em JSON, agora com o Status Geral correto
    with open(caminho_json, "w", encoding="utf-8") as f:
        json.dump(lista_logs, f, ensure_ascii=False, indent=2)
    print(f"[LOG] Log detalhado da execução salvo em: {caminho_json}")

    # Chama a função para gerar o relatório sumarizado em CSV
    # (Esta função já deve estar atualizada conforme o Passo 1)
    gerar_relatorio_sumarizado(lista_logs, timestamp)
    
    # --- INCLUSÃO (Passo 3) ---
    # Chama a nova função para gerar o arquivo Excel de reprocessamento
    # (Esta função já deve estar definida conforme o Passo 2)
    gerar_arquivo_reprocessamento(lista_logs, df_original, timestamp)
    # --- FIM DA INCLUSÃO ---

def main():
    start_time = time.time()
    print("[INÍCIO] INÍCIO DO SCRIPT RPA", flush=True)

    parser = argparse.ArgumentParser(description='RPA para atualizar processos no Legal One em lote.')
    parser.add_argument('--start', type=int, help='Número da linha (Excel) inicial para processar.')
    parser.add_argument('--end', type=int, help='Número da linha (Excel) final para processar.')
    args = parser.parse_args()
    print(f"[INFO] Argumentos recebidos: --start={args.start}, --end={args.end}", flush=True)

    usuario = os.environ.get("LEGALONE_USUARIO")
    senha = os.environ.get("LEGALONE_SENHA")
    print(f"DEBUG rpa_script: Usuário lido do ambiente: {usuario}")
    print(f"DEBUG rpa_script: Senha lida do ambiente: {'*' * len(senha) if senha else None}")
    if not usuario or not senha: print("[ERRO CRÍTICO] Variáveis de ambiente não definidas."); return

    BASE_DIR = Path(__file__).resolve().parent
    DATA_FILE = BASE_DIR / "data" / "entrada.xlsx"
    JSON_FILE = BASE_DIR / "data" / "campos_mapeados.json"
    try:
        df_full = pd.read_excel(DATA_FILE, sheet_name="Dados")
        with open(JSON_FILE, encoding="utf-8") as f: todos_campos = json.load(f)
        mapeamento = todos_campos
    except FileNotFoundError: print(f"[ERRO CRÍTICO] Arquivo não encontrado: {DATA_FILE} ou {JSON_FILE}"); return
    except Exception as e: print(f"[ERRO CRÍTICO] Erro ao carregar dados iniciais: {e}"); return

    # Lógica de Fatiamento (Mantida)
    total_rows_in_file = len(df_full)
    start_index = args.start - 2 if args.start and args.start > 1 else 0
    end_index = args.end - 2 if args.end else total_rows_in_file - 1
    if args.start is not None and args.end is not None and args.end < args.start: end_index = start_index - 1
    start_index = max(0, start_index); end_index = min(total_rows_in_file - 1, end_index)

    df_lote = pd.DataFrame()
    if start_index > end_index or start_index >= total_rows_in_file:
         excel_start_log = args.start if args.start else 2; excel_end_log = args.end if args.end else total_rows_in_file + 1
         print(f"[INFO] Lote inválido/vazio (Excel {excel_start_log}-{excel_end_log}, Índices {start_index}-{end_index}).")
    else:
        df_lote = df_full.iloc[start_index : end_index + 1].copy()
        excel_start_log = start_index + 2; excel_end_log = end_index + 2
        print(f"[INFO] Processando Lote - Linhas Excel: {excel_start_log} a {excel_end_log} (Total: {len(df_lote)})")

    if df_lote.empty:
        print("[INFO] Lote vazio. Encerrando."); salvar_log_execucao([]); return

    logs_processos = []
    browser = None; page = None; context = None; playwright_context = None

    try:
        playwright_context = sync_playwright().start()
        try:
            browser = playwright_context.chromium.launch(headless=False, timeout=90000, channel="chrome")
            context = browser.new_context(viewport={"width": 1366, "height": 768})
            context.set_default_timeout(45000)
            page = context.new_page()
            print("[INFO] Navegador e página iniciados.")
        except Exception as e: print(f"[ERRO CRÍTICO] Falha ao iniciar navegador: {e}"); return

        login_info = login_legalone(page, usuario, senha)
        print(f"[{login_info['etapa']}] {login_info['status']} ({login_info['duracao']}s)")
        if login_info["status"] == "Falha": print(f"[ERRO CRÍTICO] Erro no login: {login_info.get('mensagem', '')}"); return

        # Loop principal
        for i, row in df_lote.iterrows():
            excel_row_num = i + 2
            dados_linha = row.dropna().to_dict()
            numero = str(dados_linha.get("Processo", "")).strip()
            log = iniciar_log_processo(numero)
            log["Linha Excel"] = excel_row_num

            if not numero:
                print(f"[AVISO] Linha Excel {excel_row_num}: Sem número de processo, pulando.")
                adicionar_evento(log, "Geral", "aviso", "Validação", "Proc ausente", 0); logs_processos.append(log); continue

            print(f"\n[INFO] {datetime.now().strftime('%H:%M:%S')} - Atualizando proc: {numero} (Linha {excel_row_num})")

            # <<< Try principal para o processo >>>
            try:
                if not page or page.is_closed():
                    print("[ERRO GRAVE] Página fechada. Encerrando lote."); adicionar_evento(log, "Geral", "erro", "Conexão", "Página fechada", 0)
                    if not any(l.get("Processo") == numero and l.get("Linha Excel") == excel_row_num for l in logs_processos): logs_processos.append(log)
                    break

                # Acesso, Painéis, Preenchimentos...
                acesso = acessar_processo_para_edicao(page, numero); print(f"[{acesso['etapa']}] {acesso['status']} ({acesso['duracao']}s)"); adicionar_evento(log, acesso["etapa"], acesso["status"].lower(), "Abertura", acesso.get("mensagem", ""), acesso["duracao"])
                if acesso["status"] == "Falha": raise Exception(f"Falha Acesso: {acesso.get('mensagem', 'Erro')}")
                expandir = expandir_paineis(page); adicionar_evento(log, expandir["etapa"], expandir["status"].lower(), "Expansão", expandir.get("mensagem", ""), expandir["duracao"])
                cascata = preencher_lookups_em_cascata(page, dados_linha, mapeamento); adicionar_evento(log, cascata["etapa"], cascata["status"].lower(), "Lookups", cascata.get("mensagem", ""), cascata["duracao"])
                envolvidos = preencher_outros_envolvidos(page, dados_linha); adicionar_evento(log, envolvidos["etapa"], envolvidos["status"].lower(), "Preenchimento", envolvidos.get("mensagem", ""), envolvidos["duracao"])
                objetos = preencher_objetos(page, dados_linha); adicionar_evento(log, objetos["etapa"], objetos["status"].lower(), "Preenchimento", objetos.get("mensagem", ""), objetos["duracao"])
                pedidos = preencher_pedidos(page, dados_linha, mapeamento); adicionar_evento(log, pedidos["etapa"], pedidos["status"].lower(), "Preenchimento", pedidos.get("mensagem", ""), pedidos["duracao"])
                if dados_linha.get("Adverso Principal"): adverso = preencher_adverso_principal(page, dados_linha); adicionar_evento(log, adverso["etapa"], adverso["status"].lower(), "Principal", adverso.get("mensagem", ""), adverso["duracao"])
                else: adicionar_evento(log, "Adverso Principal", "sucesso", "Principal", "Vazio", 0)
                campos = atualizar_campos_usando_mapeamento(page, dados_linha, mapeamento); adicionar_evento(log, campos["etapa"], campos["status"].lower(), "Atualização", campos.get("mensagem", ""), campos["duracao"])

                
                # --- INCLUSÃO DA FUNÇÃO (EXATAMENTE COMO PEDIDO) ---
                inicio_negociacao = time.time()
                
                # Assumindo que a coluna no Excel se chama 'Negociacao Honorario'
                # (Se o nome da coluna for outro, troque a string "Negociacao Honorario" abaixo)
                valor_negociacao = dados_linha.get("Negociação do Contrato de Honorários")
                
                if valor_negociacao:
                    # Chama a função exatamente como foi definida (ela deve estar no escopo global)
                    # O MAPA_NEGOCIACAO_HONORARIO também deve estar no escopo global
                    sucesso_negociacao = preencher_negociacao_honorario(page, valor_negociacao, MAPA_NEGOCIACAO_HONORARIO)
                    
                    # Log manual baseado no retorno booleano da sua função
                    duracao_negociacao = round(time.time() - inicio_negociacao, 2)
                    if sucesso_negociacao:
                        adicionar_evento(log, "Negociação Honorário", "sucesso", "Preenchimento", "", duracao_negociacao)
                    else:
                        # A própria função já loga o erro no console, aqui logamos no log JSON
                        adicionar_evento(log, "Negociação Honorário", "erro", "Preenchimento", f"Falha ao preencher '{valor_negociacao}'", duracao_negociacao)
                else:
                    # Se não houver valor na planilha, loga como 'vazio'
                    adicionar_evento(log, "Negociação Honorário", "sucesso", "Preenchimento", "Vazio", 0)
                # --- FIM DA INCLUSÃO ---

                # --- INCLUSÃO CENTRO DE CUSTO ---
                inicio_centro_custo = time.time()
                
                # Pega o valor da coluna 'Centro de Custo' (conforme você informou)
                valor_centro_custo = dados_linha.get("Centro de Custo")
                
                if valor_centro_custo:
                    # ATENÇÃO: Assumindo que seu mapa global se chama 'MAPA_CENTRO_CUSTO'
                    # (Se o nome for outro, ajuste a variável aqui)
                    sucesso_cc = preencher_centro_custo(page, valor_centro_custo, MAPA_CENTRO_CUSTO)
                    
                    # Log manual baseado no retorno booleano da sua função
                    duracao_cc = round(time.time() - inicio_centro_custo, 2)
                    if sucesso_cc:
                        adicionar_evento(log, "Centro de Custo", "sucesso", "Preenchimento", "", duracao_cc)
                    else:
                        adicionar_evento(log, "Centro de Custo", "erro", "Preenchimento", f"Falha ao preencher '{valor_centro_custo}'", duracao_cc)
                else:
                    # Se não houver valor na planilha, loga como 'vazio'
                    adicionar_evento(log, "Centro de Custo", "sucesso", "Preenchimento", "Vazio", 0)
                # --- FIM DA INCLUSÃO ---

                
                # --- Salvar (BLOCO NOVO E CORRIGIDO COM LOOP) ---
                try: # Try interno para salvar
                    inicio_salvar = time.time()
                    botao_salvar = page.locator('button[name="ButtonSave"]')

                    if botao_salvar.is_visible() and botao_salvar.is_enabled():
                        print("     - Botão 'Salvar' OK. Clicando...")
                        botao_salvar.scroll_into_view_if_needed()
                        botao_salvar.click(timeout=15000)

                        # --- INÍCIO DA MODIFICAÇÃO (LÓGICA DE LOOP) ---
                        
                        try: # Try interno para o loop de espera
                            print("     - Aguardando eventos pós-salvamento (popup OU sucesso)...")

                            # 1. Definimos os localizadores para os DOIS possíveis eventos
                            popup_locator = page.locator('input#popup_ok')
                            
                            success_regex = re.compile(r"alterado com sucesso\.", re.IGNORECASE)
                            success_locator = page.locator('div.top-message-body.sucesso div.message-content').filter(has_text=success_regex)

                            # 2. Criamos um localizador "OU" (OR)
                            combined_locator = popup_locator.or_(success_locator)

                            # 3. Definimos um timeout TOTAL para toda a operação de salvar
                            max_save_time_ms = 35000  # 35 segundos no total
                            start_wait_time = time.time()

                            # 4. INICIAMOS O LOOP
                            while True:
                                # 4.1. Verificamos se o tempo total estourou
                                elapsed_ms = (time.time() - start_wait_time) * 1000
                                if elapsed_ms >= max_save_time_ms:
                                    raise PlaywrightTimeoutError(f"Timeout total de {max_save_time_ms}ms excedido esperando por popups ou sucesso.")
                                
                                # 4.2. Calculamos o tempo restante para esta iteração
                                remaining_timeout = max_save_time_ms - elapsed_ms

                                # 5. Esperamos pelo próximo evento (popup OU sucesso)
                                # Usamos .first para pegar o primeiro que aparecer
                                combined_locator.first.wait_for(state='visible', timeout=remaining_timeout)
                                
                                print("     - Evento pós-salvamento detectado. Verificando qual...")

                                # 6. Verificamos o que apareceu (sem espera)
                                if success_locator.is_visible():
                                    # CENÁRIO B: SUCESSO!
                                    print("     - Notificação de sucesso detectada. Saindo do loop.")
                                    break # Sai do loop 'while True'

                                elif popup_locator.is_visible():
                                    # CENÁRIO A: POPUP! (Pode ser o 1º, 2º, 3º...)
                                    try:
                                        # Tenta ler a mensagem para log
                                        msg_popup = page.locator("#popup_message").text_content(timeout=500)
                                        print(f"     - Popup de confirmação detectado: '{msg_popup[:50]}...'. Clicando...")
                                    except Exception:
                                        print("     - Popup de confirmação (sem msg) detectado. Clicando...")
                                    
                                    popup_locator.click()
                                    
                                    # 7. CRÍTICO: Esperar o popup desaparecer ANTES de continuar o loop
                                    print("     - Aguardando popup desaparecer...")
                                    popup_locator.wait_for(state='hidden', timeout=5000) # Espera o popup atual sumir
                                    print("     - Popup tratado. Voltando a aguardar (próximo popup ou sucesso)...")
                                    # O 'while True' continua
                                
                                else:
                                    # Segurança: se o wait_for() retornar mas nenhum for visível
                                    # (improvável, mas bom ter)
                                    print("     - [AVISO] Wait_for retornou, mas nenhum localizador visível. Tentando novamente...")
                                    page.wait_for_timeout(250) # Pequena pausa

                            # 8. Se chegamos aqui (fora do loop), o sucesso foi alcançado
                            duracao = round(time.time() - inicio_salvar, 2)
                            adicionar_evento(log, "Salvar", "sucesso", "Confirmação", "", duracao)
                            print(f"[SALVO] Processo {numero} salvo com sucesso ({duracao}s)")

                        except PlaywrightTimeoutError as e_timeout_combined:
                            # Se o timeout total estourar
                            duracao = round(time.time() - inicio_salvar, 2)
                            print(f"     - DEBUG: Timeout ({str(e_timeout_combined).splitlines()[0]}) ao esperar popup OU sucesso.")
                            mensagem = f"Falha Salvar: Nem popup nem notificação de sucesso encontrados ({duracao}s)."
                            adicionar_evento(log, "Salvar", "erro", "Confirmação", mensagem, duracao)
                            print(f"[ERRO] {mensagem} no proc {numero}")
                            if page.is_closed(): 
                                raise Exception("Página fechou durante espera.")
                            else: 
                                raise Exception(mensagem)

                        except Exception as e_combined:
                            # Tratamento de outros erros no loop
                            duracao = round(time.time() - inicio_salvar, 2)
                            print(f"     - DEBUG: Erro inesperado no loop: {type(e_combined).__name__} - {e_combined}")
                            mensagem = f"Falha Salvar: Erro ao verificar notificação/popup: {e_combined}"
                            adicionar_evento(log, "Salvar", "erro", "Confirmação", mensagem, duracao)
                            print(f"[ERRO] {mensagem} no proc {numero}")
                            if page.is_closed(): 
                                raise Exception("Página fechou durante espera.")
                            else: 
                                raise Exception(mensagem)
                        
                        # --- FIM DA MODIFICAÇÃO ---

                    else:
                        msg_botao = "Botão 'Salvar' não encontrado/visível/habilitado"
                        adicionar_evento(log, "Salvar", "erro", "Confirmação", msg_botao, 0)
                        print(f"[AVISO] {msg_botao} para o proc {numero}")
                        raise Exception(msg_botao)

                # Except alinhado com o 'try' da operação de salvar
                except Exception as e_save:
                    duracao = round(time.time() - inicio_salvar, 2) if 'inicio_salvar' in locals() else 0
                    if not any(err.get("etapa") == "Salvar" for err in log.get("Erros",[])):
                        adicionar_evento(log, "Salvar", "erro", "Confirmação", f"Erro crítico: {str(e_save)}", duracao)
                    print(f"[ERRO] Falha crítica salvar {numero}: {e_save}")
                    raise Exception(f"Erro crítico ao salvar: {e_save}") # Propaga

            # <<< Except alinhado com o try principal do processo >>>
            except Exception as e_processo:
                print(f"[ERRO INESPERADO] proc {numero} (Linha {excel_row_num}): {e_processo}")
                if not any(err.get("etapa") == "Processamento Geral" for err in log.get("Erros", [])):
                       adicionar_evento(log, "Processamento Geral", "erro", "Execução", str(e_processo), 0)

                error_msg_lower = str(e_processo).lower()
                if "target page" in error_msg_lower or "browser has been closed" in error_msg_lower or \
                   "net::err_aborted" in error_msg_lower or "browser has crashed" in error_msg_lower or \
                   "playwright playwright" in error_msg_lower:
                    print("[ERRO GRAVE] Conexão/Navegador perdido. Encerrando lote.")
                    if not any(l.get("Processo") == numero and l.get("Linha Excel") == excel_row_num for l in logs_processos): logs_processos.append(log)
                    break
                else:
                    try:
                        print("     - Tentando voltar para busca...");
                        if page and not page.is_closed():
                             page.goto("https://mdradvocacia.novajus.com.br/processos/processos/search", wait_until="load", timeout=30000)
                             print("     - Recuperação OK.")
                        else:
                            print("[ERRO GRAVE] Página fechada na recuperação. Encerrando.")
                            if not any(l.get("Processo") == numero and l.get("Linha Excel") == excel_row_num for l in logs_processos): logs_processos.append(log)
                            break
                    except Exception as e_nav:
                        print(f"[ERRO GRAVE] Falha ao voltar para busca: {e_nav}. Encerrando.")
                        adicionar_evento(log, "Recuperação", "erro", "Navegação", f"Falha: {e_nav}", 0)
                        if not any(l.get("Processo") == numero and l.get("Linha Excel") == excel_row_num for l in logs_processos): logs_processos.append(log)
                        break

            # <<< Finally alinhado com o try principal do processo >>>
            finally:
                if not any(l.get("Processo") == numero and l.get("Linha Excel") == excel_row_num for l in logs_processos):
                    logs_processos.append(log)
                print(f"[INFO] Processamento linha {excel_row_num} concluído.")

        # Fim do loop for
        print(f"\n[INFO] {datetime.now().strftime('%H:%M:%S')} - Fim do processamento do lote.")

    # Except GERAL fora do loop
    except Exception as e_global:
        print(f"[ERRO GERAL FORA DO LOOP] {e_global}")
        if not logs_processos:
            log_erro_geral = { "Processo": "ERRO_GERAL_SCRIPT", "Linha Excel": "N/A", "Status Geral": "Falha", "Erros": [{"etapa": "Inicialização", "item": "Erro Script", "mensagem": str(e_global)}], "Sucessos": [], "Duracao por etapa (s)": {} }
            logs_processos.append(log_erro_geral)

    # Finally GERAL
    finally:
        if page and not page.is_closed():
            try:
                page.close()
                print("[INFO] Página fechada.")
            except Exception as e:
                print(f"[AVISO] Erro ao fechar página: {e}") 

        if context:
            try:
                context.close()
                print("[INFO] Contexto fechado.")
            except Exception as e:
                print(f"[AVISO] Erro ao fechar contexto: {e}") 

        if browser and browser.is_connected():
            try:
                browser.close()
                print("[INFO] Navegador fechado.")
            except Exception as e:
                print(f"[AVISO] Erro ao fechar navegador: {e}") 
        elif browser:
            print("[INFO] Navegador já estava desconectado.")
        else:
            print("[INFO] Nenhuma instância de navegador foi criada.")

        if playwright_context:
            try:
                playwright_context.stop()
                print("[INFO] Playwright parado.")
            except Exception as e:
                print(f"[AVISO] Erro ao parar Playwright: {e}") 

        # Salvar logs
        print("\n[INFO] Finalizando script. Salvando log...")
        if 'salvar_log_execucao' in globals(): salvar_log_execucao(logs_processos, df_full)
        else: print("[ERRO CRÍTICO] Função 'salvar_log_execucao' não definida!")
        tempo_total = time.time() - start_time
        print(f"\n[TEMPO] Tempo total: {tempo_total:.2f} segundos.")

if __name__ == "__main__":
    # Configuração de encoding (Mantido)
    if sys.stdout.isatty() and os.name == 'nt':
        try: sys.stdout.reconfigure(encoding='utf-8'); sys.stderr.reconfigure(encoding='utf-8')
        except Exception as e: print(f"[AVISO] Falha ao reconfigurar encoding: {e}")
    main()