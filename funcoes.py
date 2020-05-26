import pandas as pd
from tqdm import tqdm

def _ponto_por_virgula(df, coluna):
    df[coluna] = df[coluna].astype(str)
    for i in df.index:
        # Trocar pontos por virgulas na coluna passada
        df.at[i, coluna] =(df.at[i, coluna].replace(".", ","))

    return df

def _pegar_tributacao(produto, ncm, app):
    """Pegar o tipo de tributacao do produto no banco de dados"""
    if pd.isna(produto) == False:
        df = pd.read_csv("banco_de_dados_icms.csv", encoding="latin1", sep=";")
        i = df[(df['Descrição do Produto'] == produto)].index
        try:
            return str(df.at[i[0], "Tipo de Tributação"])
        except:
            tipo_tributacao = app.textBox("Adicionar produto", "O produto "+str(produto)+", de NCM "+str(ncm)+", não está "
                        "cadastrado no banco de dados. Entre com o seu tipo de tributação para "
                        "cadastrá-lo.", parent=None)
            tipo_tributacao = str(tipo_tributacao).upper()

            tam = len(df["Descrição do Produto"])
            df.at[tam, "NCM"] = str(ncm)
            df.at[tam, "Tipo de Tributação"] = tipo_tributacao
            df.at[tam, "Descrição do Produto"] = produto
            df.to_csv("banco_de_dados_icms.csv", encoding='latin-1', sep=';', index=False)

            return tipo_tributacao

def _tratar_relatorio(df, app):
    """Trata o relatório cru gerado pela domínio"""# Excluir Saidas e MG e CFOP=2949
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

    df.drop(df.tail(2).index,inplace=True) # Dropa duas últimas linhas por causa do formato do relatório gerado pela Domínio
    df['CFOP'] = df['CFOP'].astype(int)
    df['Documento'] = df['Documento'].astype(int)
    df['Quantidade'] = df['Quantidade'].astype(int)

    for i in df.index:
        ncm = df.at[i, "NCM"]
        df.at[i, "Tipo de Tributação"] = str(_pegar_tributacao(df.at[i, "Descrição item"], ncm, app))

    return df

def _gerar_resumo(df, path_to_save):
    df_resumo = pd.DataFrame(columns=['NF', 'Data Entrada/Saída', 'BC USO', 'ST'])
    df_resumo['BC USO'] = 0.0
    df_resumo['ST'] = 0.0
    print ("\nGerando resumo...")
    i = path_to_save.rfind("/") # rfind pega a ultima ocorrência
    path_to_save = path_to_save[:i]
    path_to_save += "/resumo.csv"

    documento_antes = ""

    for i in tqdm(df.index):
        documento = str(df.at[i, 'Documento'])

        if documento == documento_antes:
            if df.at[i, "Tipo de Tributação"] == "BC USO":
                df_resumo.at[j, 'BC USO'] += float(df.at[i, "Valor ICMS"])

            else:
                df_resumo.at[j, 'ST'] += float(df.at[i, "Valor ICMS"])
        else:
            j += 1
            df_resumo.at[j, 'NF'] = documento
            df_resumo.at[j, 'Data Entrada/Saída'] = df.at[i, 'Data Entrada/Saída']
            if df.at[i, "Tipo de Tributação"] == "BC USO":
                df_resumo.at[j, 'BC USO'] = float(df.at[i, "Valor ICMS"])
                df_resumo.at[j, 'ST'] = 0.0
            else:
                df_resumo.at[j, 'BC USO'] = 0.0
                df_resumo.at[j, 'ST'] = float(df.at[i, "Valor ICMS"])
            documento_antes = documento



    df_resumo = _ponto_por_virgula(df_resumo, 'BC USO')
    df_resumo = _ponto_por_virgula(df_resumo, 'ST')
    df_resumo.to_csv(path_to_save, encoding='latin-1', sep=';', index=False)

def _calcular(nome_arquivo, path_to_save, app):
    df = pd.read_excel(nome_arquivo[0], encoding='latin-1', sep=';')
    df = _tratar_relatorio(df, app)
    df['Valor ICMS'] = 0.0

    aliq_interna = float(app.getEntry("aliq_interna_entry"))
    aliq_interestadual = app.getEntry("aliq_interestadual_entry")

    for i in tqdm(df.index):
        if str(df.at[i, "CFOP"]) != "2949":
            if (df.at[i, "Tipo de Tributação"] == "BC USO"):
                AB = float(df.at[i, "Valor total"] + df.at[i, "Valor do IPI"] + df.at[i, "Despesas Aces."] - df.at[i, "Descontos"])
                AC = float(aliq_interestadual)
                AD = float(AB * AC)
                AE = float(AB * (1 - AC))
                AF = float(aliq_interna)
                AG = float(AE / (1 - AF))
                AH = float(AG * AF)
                AI = float(AH - AD)
                if float(AI) >= 0.01:
                    df.at[i, 'Valor ICMS'] = round(float(AI), 2)
            elif (df.at[i, "Tipo de Tributação"] == "ST USO"):
                AB = float(df.at[i, "Valor total"] + df.at[i, "Valor do IPI"] + df.at[i, "Despesas Aces."] + df.at[i, "Frete"])
                AC = float(aliq_interestadual)
                AD = float(AB * AC)
                AE = float(AB * (1 - AC))
                AF = float(aliq_interna)
                AG = float(AE / (1 - AF))
                AH = float(AG * AF)
                AI = float(AH - AD)
                if float(AI) >= 0.01:
                    df.at[i, 'Valor ICMS'] = round(float(AI), 2)

            elif (df.at[i, "Tipo de Tributação"] == "BC") or (df.at[i, "Tipo de Tributação"] == 'None'): # BC faz nada
                pass

            else:
                AA = float(df.at[i, "Tipo de Tributação"])
                AB = float(df.at[i, "Valor total"] + df.at[i, "Frete"] + df.at[i, "Despesas Aces."])
                AC = float(df.at[i, "Valor do IPI"])
                AD = float(aliq_interestadual)
                AE = float(aliq_interna)
                AF = float(AA)
                AG = float(AA)
                AH = float(df.at[i, "Valor do ST"])
                resultado = float((((((AB + AC) * (1 + AG)) * AE) - (AB * AD)) - AH))
                if resultado >= 0.01:
                    df.at[i, 'Valor ICMS'] = round(resultado, 2)

    _gerar_resumo(df, path_to_save)
    df = _ponto_por_virgula(df, 'Valor ICMS')
    df.to_csv(path_to_save, encoding='latin-1', sep=';', index=False)
    app.infoBox("Arquivo Salvo", "O relatório com o valor de ICMS foi salvo.")