import pandas as pd
import pywhatkit
import time
from datetime import datetime, time as dt_time


# =================== CONFIGURAÇÕES ===================
ARQUIVO = "ACAI_AUTO.xlsx"  # nome do seu arquivo
HORA_INICIO = 8
HORA_FIM = 17

# Mensagens baseadas no tipo de negócio
MENSAGENS = {
    "Açaíteria": "Olá {nome}! Aqui é da fábrica de polpas naturais, estamos com condições especiais para a sua açaíteria!",
    "Casa de Sucos": "Oi {nome}! Trabalhamos com polpas naturais ideais para casas de sucos como a sua!",
    "Restaurante": "Olá {nome}, temos polpas congeladas perfeitas para o seu restaurante!",
    "Supermercado": "Oi {nome}! Você sabia que polpas naturais são ótimas para vender no seu supermercado?",
    "Catering": "Olá {nome}, temos soluções de polpas ideais para o seu serviço de catering.",
    "Indústria de Alimentos": "Olá {nome}, trabalhamos com fornecimento de polpas em grande escala para indústrias como a sua!",
}

# =================== FUNÇÕES ===================

def dentro_do_horario():
    agora = datetime.now().time()
    return dt_time(HORA_INICIO) <= agora <= dt_time(HORA_FIM)


def limpar_telefone(numero):
    import re
    # Remove tudo que não for número
    numeros = re.findall(r'\d+', numero)
    numero_limpo = ''.join(numeros)
    # Adiciona o +55 no começo
    if numero_limpo.startswith("55"):
        return "+" + numero_limpo
    elif numero_limpo.startswith("0"):
        return "+55" + numero_limpo[1:]
    elif len(numero_limpo) >= 10:
        return "+55" + numero_limpo
    else:
        return None

def gerar_mensagem(nome, tipo):
    tipo = tipo.strip()
    base = MENSAGENS.get(tipo, "Olá {nome}, temos ofertas em polpas naturais que podem interessar para o seu negócio!")
    return base.format(nome=nome)

# =================== EXECUÇÃO ===================

# Lê a planilha
df = pd.read_excel(ARQUIVO)
df.columns = ["Nome", "Tipo", "Telefone"]
df = df[1:]  # remove o cabeçalho extra
df = df.dropna(subset=["Nome", "Tipo", "Telefone"])

print("Iniciando o Robô do Açaí...")

while True:
    if not dentro_do_horario():
        print("Fora do horário comercial. Encerrando robô.")
        break

    linha = df.iloc[0]  # linha de teste
    nome = linha["Nome"]
    tipo = linha["Tipo"]
    telefones_raw = str(linha["Telefone"]).split(",")

    enviado = False
    for t in telefones_raw:
        telefone_formatado = limpar_telefone(t)
        if telefone_formatado:
            mensagem = gerar_mensagem(nome, tipo)
            print(f"Enviando para {telefone_formatado}: {mensagem}")
            pywhatkit.sendwhatmsg_instantly(telefone_formatado, mensagem, wait_time=15)
            enviado = True
            break  # só um teste, envia para o primeiro válido

    if not enviado:
        print("Não foi possível enviar mensagem. Nenhum número válido.")

    break  # envia só uma vez no teste
