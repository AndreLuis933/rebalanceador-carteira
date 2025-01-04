import pandas as pd
from bs4 import BeautifulSoup
import pyperclip
import time
import os

try:
    TOTAL = float(input("Digite o valor total: "))
    LIMITE_DE_ENTRADA = float(input("Digite o limite de entrada: "))
except ValueError:
    raise ValueError('Digite um número válido para "TOTAL" e "LIMITE_DE_ENTRADA"')
    

def obter_conteudo_area_transferencia(timeout=1800):
    """Obtém o conteúdo da área de transferência. Lança um erro se não estiver acessível ou vazia."""
    start_time = time.time()
    while time.time() - start_time < timeout:
        try:
            html_content = pyperclip.paste().strip()
            if "<body>" in html_content:
                return html_content
            time.sleep(0.5)
        except Exception as e:
            print(f"Erro ao acessar a área de transferência: {e}", "ERRO")
    else:
        print("Tempo limite atingido. Nenhum conteúdo capturado.")


html_content = obter_conteudo_area_transferencia()


def extrair_dados():
    soup = BeautifulSoup(html_content, "html.parser")
    todas_categorias = soup.find_all(class_="accordion")
    dataframes = []

    for categoria in todas_categorias:
        if categoria.text:
            cards = categoria.find_all(class_="card-body")

            # Coletando os dados de cada classe de ativos
            dados = [
                {
                    "Nome": card.find(class_="text-muted").text,
                    "Porcentagem": card.find("div", class_="porcent").text,
                    "Representaçao": card.find(class_="progress-bar")["style"].split(
                        ":"
                    )[1],
                }
                for card in cards
            ]

            # Criando DataFrame para a classe atual
            df = pd.DataFrame(dados)
            dataframes.append(df)

    return dataframes


def tratamento_dataframes(df, total):

    df[["Atual", "Meta"]] = (
        df["Porcentagem"].str.extract(r"Atual (.*?)%Meta (.*?)%").astype(float)
    )
    df["Representaçao"] = df["Representaçao"].str.rstrip("%;").astype(float)
    df = df.drop("Porcentagem", axis=1)
    df = df[["Nome", "Atual", "Meta", "Representaçao"]]
    df["Total"] = total
    df["Valor_atual"] = round(df["Atual"] * df["Total"] / 100, 2)
    df["Esperado"] = round(df["Meta"] * df["Total"] / 100, 2)
    df["Movimentacao"] = round(df["Valor_atual"] - df["Esperado"], 2)
    return df


def limitar_movimentacao(df, limite):
    # Verifica se a soma das movimentações ultrapassa o limite
    soma_movimentacao = df["Movimentacao"].sum()
    if abs(soma_movimentacao) > limite:
        # Ajusta proporcionalmente os valores de movimentação para atender ao limite
        fator_ajuste = limite / abs(soma_movimentacao)
        df["Movimentacao"] = round(df["Movimentacao"] * fator_ajuste, 2)
    return df


dataframes = extrair_dados()

for data in dataframes:

    os.makedirs("Percentual", exist_ok=True)

    df = tratamento_dataframes(data, TOTAL)
    df.loc[0, ["Movimentacao", "Valor_atual", "Esperado", "Total"]] = 0
    #  se quiser limitar a quantidade a ser movimentada
    df = limitar_movimentacao(df, LIMITE_DE_ENTRADA)
    
    df.to_excel(rf"Percentual\{df['Nome'][0]}.xlsx", index=False)
