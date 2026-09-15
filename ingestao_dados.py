import os
import sqlite3
import pandas as pd


def processar_e_padronizar():
    caminho_base = "."
    anos_pastas = ["2020", "2021", "2022", "2023", "2024"]

    dados_aba1 = []
    dados_aba3 = []

    print("🚀 Processando arquivos (.xlsx / .xlsm)...")

    for ano in anos_pastas:
        caminho_ano = os.path.join(caminho_base, ano)
        if not os.path.exists(caminho_ano):
            continue

        arquivos = [
            f
            for f in os.listdir(caminho_ano)
            if f.lower().endswith((".xlsx", ".xlsm", ".xls"))
            and not f.startswith("~$")
        ]
        print(f"\n📁 Pasta {ano}: {len(arquivos)} arquivo(s) encontrado(s).")

        for arquivo in arquivos:
            caminho_arquivo = os.path.join(caminho_ano, arquivo)
            print(f"  📂 Lendo: {ano}/{arquivo}")

            nome_sem_ext = os.path.splitext(arquivo)[0]
            mes_extraido = (
                nome_sem_ext.split()[0].upper()
                if nome_sem_ext.split()
                else "N/A"
            )

            try:
                excel_file = pd.ExcelFile(caminho_arquivo, engine="openpyxl")
                nomes_abas = excel_file.sheet_names

                # Aba 1 (ESTOQUE PA)
                df_1 = pd.read_excel(
                    excel_file, sheet_name=0, header=1, engine="openpyxl"
                )
                df_1["ANO_ORIGEM"] = ano
                df_1["MES_ORIGEM"] = mes_extraido
                dados_aba1.append(df_1)

                # Aba 3 (SAÍDAS)
                idx_aba3 = 2 if len(nomes_abas) >= 3 else (len(nomes_abas) - 1)
                df_3 = pd.read_excel(
                    excel_file,
                    sheet_name=idx_aba3,
                    header=1,
                    engine="openpyxl",
                )
                df_3["ANO_ORIGEM"] = ano
                df_3["MES_ORIGEM"] = mes_extraido
                dados_aba3.append(df_3)

            except Exception as e:
                print(f"   ❌ Erro ao ler {arquivo}: {e}")

    conn = sqlite3.connect("estoque.db")

    # --- ABA 1 (CADASTRO PRODUTOS) ---
    if dados_aba1:
        df_final1 = pd.concat(dados_aba1, ignore_index=True)
        df_final1.columns = [str(c).strip() for c in df_final1.columns]

        colunas_validas1 = [
            c
            for c in df_final1.columns
            if c and not c.startswith("Unnamed") and c.upper() != "NONE"
        ]
        df_final1 = df_final1[colunas_validas1]
        df_final1 = df_final1.loc[:, ~df_final1.columns.duplicated()]

        # Converte qualquer coluna de data/Timestamp para texto (evita erro no Python 3.14)
        for col in df_final1.select_dtypes(
            include=["datetime64", "datetimetz"]
        ).columns:
            df_final1[col] = df_final1[col].astype(str)

        df_final1.to_sql(
            "cadastro_produtos", conn, if_exists="replace", index=False
        )
        print(
            f"\n📦 Tabela 'cadastro_produtos' salva com {len(df_final1)} registros."
        )

    # --- ABA 3 (MOVIMENTACAO VENDAS) ---
    if dados_aba3:
        df_final3 = pd.concat(dados_aba3, ignore_index=True)

        rename_dict = {}
        for col in df_final3.columns:
            c_str = str(col).strip().upper()
            if c_str in ["CODIGO", "CÓDIGO", "COD", "ID_CODIGO"]:
                rename_dict[col] = "ID_CODIGO"
            elif "PRODUTO" in c_str or "NOME DO PRODUTO" in c_str:
                rename_dict[col] = "NOME_DO_PRODUTO"
            elif (
                "KG" in c_str
                or "QTD" in c_str
                or "QUANTIDADE" in c_str
                or "SAIDA" in c_str
            ):
                rename_dict[col] = "QUANTIDADE_KG"
            elif "CLIENTE" in c_str and "ID" not in c_str:
                rename_dict[col] = "NOME_CLIENTE"

        df_final3.rename(columns=rename_dict, inplace=True)

        if (
            "QUANTIDADE_KG" not in df_final3.columns
            and len(df_final3.columns) >= 11
        ):
            col_k = df_final3.columns[10]
            df_final3.rename(columns={col_k: "QUANTIDADE_KG"}, inplace=True)

        if (
            "NOME_CLIENTE" not in df_final3.columns
            and len(df_final3.columns) >= 14
        ):
            col_n = df_final3.columns[13]
            df_final3.rename(columns={col_n: "NOME_CLIENTE"}, inplace=True)

        df_final3.columns = [str(c).strip() for c in df_final3.columns]
        colunas_validas3 = [
            c
            for c in df_final3.columns
            if c and not c.startswith("Unnamed") and c.upper() != "NONE"
        ]
        df_final3 = df_final3[colunas_validas3]
        df_final3 = df_final3.loc[:, ~df_final3.columns.duplicated()]

        if "QUANTIDADE_KG" in df_final3.columns:
            df_final3["QUANTIDADE_KG"] = pd.to_numeric(
                df_final3["QUANTIDADE_KG"], errors="coerce"
            ).fillna(0)
        else:
            df_final3["QUANTIDADE_KG"] = 0

        if "NOME_CLIENTE" in df_final3.columns:
            df_final3["NOME_CLIENTE"] = (
                df_final3["NOME_CLIENTE"].fillna("NÃO INFORMADO").astype(str)
            )
            df_final3["ID_CLIENTE"] = (
                df_final3["NOME_CLIENTE"].astype("category").cat.codes + 1
            )
        else:
            df_final3["ID_CLIENTE"] = 1
            df_final3["NOME_CLIENTE"] = "Cliente Geral"

        # Converte todas as colunas de data/Timestamp para texto (solução do erro)
        for col in df_final3.columns:
            if (
                pd.api.types.is_datetime64_any_dtype(df_final3[col])
                or df_final3[col]
                .apply(lambda x: isinstance(x, pd.Timestamp))
                .any()
            ):
                df_final3[col] = df_final3[col].astype(str)

        df_final3.to_sql(
            "movimentacao_vendas", conn, if_exists="replace", index=False
        )
        print(
            f"🚚 Tabela 'movimentacao_vendas' salva com {len(df_final3)} registros!"
        )

    conn.close()
    print("\n🎉 Processo concluído com sucesso!")


if __name__ == "__main__":
    processar_e_padronizar()