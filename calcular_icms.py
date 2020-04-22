import pandas as pd

def _calcular(nome_arquivo, path_to_save):

    df = pd.read_excel(nome_arquivo, encoding='latin-1', sep=';')

    # Excluir Saidas e MG
    # Remove saídas com CST ICMS diferente de 40
    indexes_to_drop = df[(df['Tipo'] == "Saída")].index
    df.drop(indexes_to_drop, inplace=True)
    indexes_to_drop = df[(df['UF'] == "MG")].index
    df.drop(indexes_to_drop, inplace=True)

    df["Data Entrada/Saída"] = df["Unnamed: 6"]
    df.drop(["Unnamed: 6"], axis=1, inplace=True)

    df["Descrição item"] = df["Unnamed: 8"]
    df.drop(["Unnamed: 8"], axis=1, inplace=True)
    df.drop(["Unnamed: 9"], axis=1, inplace=True)

    df.rename(columns={"Unidade\n": "Unidade"}, inplace=True)
    df["Unidade"] = df["Unnamed: 11"]
    df.drop(["Unnamed: 11"], axis=1, inplace=True)

    df.rename(columns={"NCM\n": "NCM"}, inplace=True)
    df["NCM"] = df["Unnamed: 13"]
    df.drop(["Unnamed: 13"], axis=1, inplace=True)

    df.rename(columns={"Quantidade\n": "Quantidade"}, inplace=True)
    df["Quantidade"] = df["Unnamed: 15"]
    df.drop(["Unnamed: 15"], axis=1, inplace=True)

    df.rename(columns={"Valor unitário\n": "Valor unitário"}, inplace=True)
    df["Valor unitário"] = df["Unnamed: 17"]
    df.drop(["Unnamed: 17"], axis=1, inplace=True)

    df.rename(columns={"Valor total\n": "Valor total"}, inplace=True)
    df["Valor total"] = df["Unnamed: 19"]
    df.drop(["Unnamed: 19"], axis=1, inplace=True)

    df.rename(columns={"Valor do IPI\n": "Valor do IPI"}, inplace=True)
    df["Valor do IPI"] = df["Unnamed: 21"]
    df.drop(["Unnamed: 21"], axis=1, inplace=True)

    df["Frete"] = df["Unnamed: 23"]
    df.drop(["Unnamed: 23"], axis=1, inplace=True)

    df["Despesas Aces."] = df["Unnamed: 25"]
    df.drop(["Unnamed: 25"], axis=1, inplace=True)

    df["Descontos"] = df["Unnamed: 27"]
    df.drop(["Unnamed: 27"], axis=1, inplace=True)

    df["Valor Contábil"] = df["Unnamed: 29"]
    df.drop(["Unnamed: 29"], axis=1, inplace=True)

    df["Base ST"] = df["Unnamed: 31"]
    df.drop(["Unnamed: 31"], axis=1, inplace=True)

    df["Aliq. ST"] = df["Unnamed: 33"]
    df.drop(["Unnamed: 33"], axis=1, inplace=True)

    df["Valor do ST"] = df["Unnamed: 35"]
    df.drop(["Unnamed: 35"], axis=1, inplace=True)

    df.rename(columns={"Base do ICMS\n": "Base do ICMS"}, inplace=True)
    df["Base do ICMS"] = df["Unnamed: 37"]
    df.drop(["Unnamed: 37"], axis=1, inplace=True)

    df.rename(columns={"Aliq. ICMS\n": "Aliq. ICMS"}, inplace=True)
    df["Aliq. ICMS"] = df["Unnamed: 39"]
    df.drop(["Unnamed: 39"], axis=1, inplace=True)

    df.rename(columns={"Valor do ICMS\n": "Valor do ICMS"}, inplace=True)
    df["Valor do ICMS"] = df["Unnamed: 41"]
    df.drop(["Unnamed: 41"], axis=1, inplace=True)

    df["CST ICMS"] = df["Unnamed: 43"]
    df.drop(["Unnamed: 43"], axis=1, inplace=True)

    print(df.columns)

    df.to_csv(path_to_save, encoding='latin-1', sep=';', index= False)

