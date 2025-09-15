import re
import os
import time
import json
import random
import logging
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor

import pandas as pd
import requests
from bs4 import BeautifulSoup
from docx import Document
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from dateutil.parser import parse

#CONFIGURAÇÕES

CONFIG = {
    'caminho_docx': "SEU_ARQUIVO.docx",
    'timeout_requisicao': 20,
    'user_agent': "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36",
    'prefixo_excluir': "https://sgp.madrix.app/",
    'max_tentativas': 3,
    'fator_backoff': 1,
    'max_threads': 4,
    'delay_entre_requisicoes': 1.5,
    'nivel_log': "INFO"
}

LINK_PATTERN = re.compile(r"https?://[^\s)'\"]+")
MESES_EM_PORTUGUES = [
    "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
]

PATTERNS_ULTIMA_MOD = [
    re.compile(r"Última modificação:.*?(\d{1,2}\s+de\s+[a-zA-Zçãõé]+\s+de\s+\d{4})", re.I),
    re.compile(r"Última modificação:.*?(\d{2}/\d{2}/\d{4})", re.I),
]
PATTERNS_ATUALIZACAO = [
    re.compile(r"Atualizad[oa]\s+(?:em|:)?\s*(\d{2}/\d{2}/\d{4})", re.I),
    re.compile(r"Atualizad[oa]\s+(?:em|:)?\s*(\d{1,2}\s+de\s+[a-zA-Zçãõé]+\s+de\s+\d{4})", re.I),
]
PATTERNS_PUBLICACAO = [
    re.compile(r"Publicado\s+(?:em|:)?\s*(\d{2}/\d{2}/\d{4})", re.I),
    re.compile(r"Publicado\s+(?:em|:)?\s*(\d{1,2}\s+de\s+[a-zA-Zçãõé]+\s+de\s+\d{4})", re.I),
]
GENERIC_PATTERNS = [
    re.compile(r"(\d{1,2}\s+de\s+[a-zA-Zçãõé]+\s+de\s+\d{4})", re.I),
    re.compile(r"(\d{2}/\d{2}/\d{4})", re.I),
    re.compile(r"(\d{4}-\d{2}-\d{2})", re.I),
]

def carregar_configuracoes():
    return CONFIG

def configurar_logging(nivel_log: str):
    logging.basicConfig(level=nivel_log.upper(), format='%(asctime)s - %(levelname)s - %(message)s', handlers=[logging.StreamHandler()])

def criar_sessao_http(config: dict) -> requests.Session:
    session = requests.Session()
    session.headers.update({"User-Agent": config['user_agent']})
    retries = Retry(total=config['max_tentativas'], backoff_factor=config['fator_backoff'], status_forcelist=[429, 500, 502, 503, 504])
    adapter = HTTPAdapter(max_retries=retries)
    session.mount("http://", adapter)
    session.mount("https://", adapter)
    return session

def extrair_links_do_docx(caminho_docx: str) -> list[tuple[str, str, str]]:
    doc, links_encontrados = Document(caminho_docx), {}
    for table in doc.tables:
        for row in table.rows:
            if len(row.cells) < 2: continue
            codigo, titulo = row.cells[0].text.strip(), row.cells[1].text.strip()
            if 'código' in codigo.lower() or 'título' in titulo.lower(): continue
            for cell in row.cells:
                for match in LINK_PATTERN.finditer(cell.text):
                    link = match.group(0).strip()
                    if link not in links_encontrados: links_encontrados[link] = (codigo, titulo)
    return [(link, codigo, titulo) for link, (codigo, titulo) in links_encontrados.items()]

def buscar_data_de_atualizacao(sessao: requests.Session, link: str, config: dict) -> tuple[str | None, str | None]:
    try:
        time.sleep(random.uniform(config['delay_entre_requisicoes'] * 0.5, config['delay_entre_requisicoes'] * 1.5))
        response = sessao.get(link, timeout=config['timeout_requisicao'])
        response.raise_for_status()
        
        soup = BeautifulSoup(response.text, "html.parser")
        for script in soup(["script", "style"]): script.decompose()
        texto_visivel = soup.get_text(" ", strip=True)

        # 1. Prioridade Máxima: "Última modificação"
        for pattern in PATTERNS_ULTIMA_MOD:
            if match := pattern.search(texto_visivel): return match.group(1).strip(), None

        # 2. Segunda Prioridade: "Atualizado"
        for pattern in PATTERNS_ATUALIZACAO:
            if match := pattern.search(texto_visivel): return match.group(1).strip(), None
        
        # 3. Terceira Prioridade: "Publicado"
        for pattern in PATTERNS_PUBLICACAO:
            if match := pattern.search(texto_visivel): return match.group(1).strip(), None

        # 4. Lógica para detectar páginas de índice
        all_generic_patterns = re.compile('|'.join(p.pattern for p in GENERIC_PATTERNS), re.I)
        if len(all_generic_patterns.findall(texto_visivel)) > 3:
            return None, None
            
        # 5. Última tentativa: Procurar por qualquer data genérica
        for pattern in GENERIC_PATTERNS:
            if match := pattern.search(texto_visivel): return match.group(0).strip(), None
        
        return None, None
        
    except requests.exceptions.RequestException as e:
        return None, f"Erro de conexão: {type(e).__name__}"
    except Exception as e:
        logging.error("Erro inesperado ao processar o link %s: %s", link, e)
        return None, "Erro inesperado durante a busca"

def verificar_link(dados: tuple) -> dict | None:
    sessao, link, codigo, titulo, config, mes_verificacao_str = dados
    if link.lower().startswith(config['prefixo_excluir']): return None
    resultado = {"Código da Norma": codigo, "Título": titulo, "Link": link, "Mês da Verificação": mes_verificacao_formatado, "Data de Atualização Encontrada": "", "Situação": ""}
    
    data_str, erro = buscar_data_de_atualizacao(sessao, link, config)
    if erro: 
        resultado["Situação"] = erro
    elif data_str:
        resultado["Data de Atualização Encontrada"] = data_str
        resultado["Situação"] = "Não atualizado"
        agora = datetime.now()
        try:
            parsed_date = parse(data_str, dayfirst=True)
            
            if parsed_date and parsed_date.year == agora.year and parsed_date.month == agora.month: 
                resultado["Situação"] = "Atualizado"
        except (ValueError, TypeError):
            logging.warning("Não foi possível analisar a data '%s' para verificação.", data_str)
            pass
    else:
        resultado["Situação"] = "Por favor, verificar atualização da norma manualmente"
    return resultado

def executar_verificacao(caminho_do_docx: str) -> tuple[pd.DataFrame, str]:
    config = carregar_configuracoes()
    configurar_logging(config['nivel_log'])
    
    agora = datetime.now()
    mes_atual_str, ano_atual_str = MESES_EM_PORTUGUES[agora.month - 1], str(agora.year)
    global mes_verificacao_formatado
    mes_verificacao_formatado = f"{mes_atual_str}/{ano_atual_str}"
    nome_arquivo_excel = f"verificacao_PQ10_{mes_atual_str.lower()}_{ano_atual_str}.xlsx"

    links_data = extrair_links_do_docx(caminho_do_docx)
    if not links_data: return pd.DataFrame(), "nenhum_link_encontrado.xlsx"

    resultados = []
    sessao = criar_sessao_http(config)
    tarefas = [(sessao, link, c, t, config, mes_verificacao_formatado) for link, c, t in links_data]

    with ThreadPoolExecutor(max_workers=config['max_threads']) as executor:
        for resultado in executor.map(verificar_link, tarefas):
            if resultado: resultados.append(resultado)

    if not resultados: return pd.DataFrame(), "nenhum_resultado.xlsx"
    
    df = pd.DataFrame(resultados)
    return df, nome_arquivo_excel