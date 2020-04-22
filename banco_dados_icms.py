import pandas as pd
from tqdm import tqdm

# df = pd.DataFrame({"NCM":[], "Descrição do Produto":[], "Tipo de Tributação":[]})
# df.to_csv("banco_de_dados_icms.csv", encoding="latin1", sep=";", index=False)

def _popular_banco(relatorios):
    base = pd.read_csv("banco_de_dados_icms.csv", encoding="latin1", sep=";")

    j = len(base["Descrição do Produto"])
    for relatorio in relatorios:
        df = pd.read_excel(relatorio, encoding='latin-1', sep=';')
        print(df.columns)
        for i in tqdm(df.index):

            descricao = df.at[i, "Descrição item"]

            if descricao not in base["Descrição do Produto"].tolist():
                base.at[j, "Descrição do Produto"] = descricao
                base.at[j, "NCM"] = df.at[i, "NCM\n"]
                base.at[j, "Tipo de Tributação"] = str(df.at[i, "Unnamed: 8"])
                j += 1
        base.to_csv("banco_de_dados_icms.csv", encoding="latin1", sep=";", index=False)



