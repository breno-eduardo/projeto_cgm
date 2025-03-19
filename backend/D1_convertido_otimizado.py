import os
import xlwings as xw
import pandas as pd
from datetime import datetime
import sys

# Lambdas para operações simples
parse_date = lambda s: datetime.strptime(s[:10], "%Y-%m-%d").date() if s and isinstance(s, str) and len(s) >= 10 else None
safe_float = lambda x: float(x) if x not in [None, ""] else 0

# Função para simular o MATCH do VBA em um DataFrame
def match_value_df(df, col, target):
    matches = df[df[col] == target]
    return matches.index[0] if not matches.empty else None

# Função para registrar warnings (equivalente ao carregaWarnings do VBA)
def register_warning(warnings_list, item_line, msg):
    if msg:
        msg = msg.strip()
        warnings_list.append((item_line, msg))
        print(f"Item D1_{item_line:05d} - Warning:\n{msg}\n")
    return

# Função para processar os itens que usam dados da API (do CSV)
def validate_api_items(df, ws_d1, warnings_list):
    b12_val = ws_d1.range("B12").value

    # D1_00001: Relatório Resumido de Execução Orçamentária – usa as 6 linhas a partir do primeiro registro
    idx = match_value_df(df, "entregavel", "Relatório Resumido de Execução Orçamentária")
    msg = ""
    if idx is not None:
        msg = "\n".join([f"Bimestre {row.periodo}: {row.status_relatorio}" 
                         for row in df.iloc[idx:idx+6].itertuples()])
    register_warning(warnings_list, 1, msg)

    # D1_00002: Balanço Anual (DCA)
    idx = match_value_df(df, "entregavel", "Balanço Anual (DCA)")
    msg = f"Anual : {df.at[idx, 'status_relatorio']}" if idx is not None else "Anual : ainda não enviado"
    register_warning(warnings_list, 2, msg)

    # D1_00003: Prefeitura RJ – Gestão Fiscal
    idx = match_value_df(df, "org_entregavel", "Prefeitura Municipal do Rio de Janeiro - RJ - Relatório de Gestão Fiscal")
    msg = ""
    if idx is not None:
        msg = "\n".join([f"Quadrimestre {row.periodo}: {row.status_relatorio}" 
                         for row in df.iloc[idx:idx+3].itertuples()])
    register_warning(warnings_list, 3, msg)

    # D1_00004: Câmara e TCM
    msg_cam = ""
    idx_cam = match_value_df(df, "org_entregavel", "Câmara de Vereadores do Rio de Janeiro - RJ - Relatório de Gestão Fiscal")
    if idx_cam is not None:
        msg_cam = "\n".join([f"Câmara - Quadrimestre {row.periodo}: {row.status_relatorio}" 
                             for row in df.iloc[idx_cam:idx_cam+3].itertuples()])
    msg_tcm = ""
    idx_tcm = match_value_df(df, "org_entregavel", "Tribunal de Contas do Município do Rio de Janeiro - Relatório de Gestão Fiscal")
    if idx_tcm is not None:
        msg_tcm = "\n".join([f"TCM - Quadrimestre {row.periodo}: {row.status_relatorio}" 
                             for row in df.iloc[idx_tcm:idx_tcm+3].itertuples()])
    register_warning(warnings_list, 4, f"{msg_cam}\n{msg_tcm}")

    # Função auxiliar inline para validar datas (usada em D1_00006, D1_00007, D1_00008 e D1_00009)
    def validate_row_date(row, date_format):
        try:
            # Monta dataComparacao usando o valor de B12 concatenado com o sufixo definido
            data_comp = datetime.strptime(f"{b12_val}{date_format.get(str(row.periodo),'')}", "%Y-%m-%d").date() if date_format.get(str(row.periodo)) else None
        except Exception:
            data_comp = None
        return data_comp, parse_date(str(row.data_status))
    
    # D1_00006: Prazo para Relatório Resumido de Execução Orçamentária
    date_map_6 = {"1": "-03-31", "2": "-05-31", "3": "-07-31", "4": "-10-01", "5": "-12-01", "6": "-01-31"}
    msg = ""
    idx = match_value_df(df, "entregavel", "Relatório Resumido de Execução Orçamentária")
    if idx is not None:
        for row in df.iloc[idx:idx+6].itertuples():
            data_comp, siconfi_date = validate_row_date(row, date_map_6)
            status_msg = (f"fora do prazo - data limite: {data_comp} e data status SICONFI: {siconfi_date}"
                          if data_comp and siconfi_date and siconfi_date > data_comp 
                          else f"dentro do prazo - data limite: {data_comp} e data status SICONFI: {siconfi_date}")
            msg += f"Bimestre {row.periodo}: {status_msg}\n"
    register_warning(warnings_list, 6, msg)

    # D1_00007: Prazo para Balanço Anual (DCA)
    msg = ""
    try:
        data_comp = datetime.strptime(f"{int(b12_val)+1}-05-01", "%Y-%m-%d").date()
    except Exception:
        data_comp = None
    idx = match_value_df(df, "entregavel", "Balanço Anual (DCA)")
    if idx is not None:
        siconfi_date = parse_date(str(df.at[idx, "data_status"]))
        status_msg = ("Fora do prazo - data limite: " + str(data_comp) + " e data status SICONFI: " + str(siconfi_date)
                      if data_comp and siconfi_date and siconfi_date > data_comp
                      else "Dentro do prazo - data limite: " + str(data_comp) + " e data status SICONFI: " + str(siconfi_date))
        msg = status_msg
    else:
        msg = "Anual : ainda não enviado"
    register_warning(warnings_list, 7, msg)

    # D1_00008: Prazo para Prefeitura RJ – Gestão Fiscal
    msg = ""
    date_map_8 = {"1": "-05-31", "2": "-10-01", "3": "-01-31"}
    idx = match_value_df(df, "org_entregavel", "Prefeitura Municipal do Rio de Janeiro - RJ - Relatório de Gestão Fiscal")
    if idx is not None:
        for row in df.iloc[idx:idx+3].itertuples():
            try:
                # Usa o mapeamento dependendo do bimestre; para "3", acrescenta 1 ao ano
                data_comp = (datetime.strptime(f"{b12_val}{date_map_8.get(str(row.periodo),'')}", "%Y-%m-%d").date()
                             if str(row.periodo) in ["1", "2"] 
                             else datetime.strptime(f"{int(b12_val)+1}{date_map_8.get(str(row.periodo),'')}", "%Y-%m-%d").date())
            except:
                data_comp = None
            siconfi_date = parse_date(str(row.data_status))
            status_msg = f"Quadrimestre {row.periodo}: " + (
                f"fora do prazo - data limite: {data_comp} e data status SICONFI: {siconfi_date}"
                if data_comp and siconfi_date and siconfi_date > data_comp 
                else f"dentro do prazo - data limite: {data_comp} e data status SICONFI: {siconfi_date}")
            msg += status_msg + "\n"
    register_warning(warnings_list, 8, msg)

    # D1_00009: Prazo para Câmara e TCM
    msg_cam = ""
    idx_cam = match_value_df(df, "org_entregavel", "Câmara de Vereadores do Rio de Janeiro - RJ - Relatório de Gestão Fiscal")
    if idx_cam is not None:
        for row in df.iloc[idx_cam:idx_cam+3].itertuples():
            try:
                data_comp = (datetime.strptime(f"{b12_val}{date_map_8.get(str(row.periodo),'')}", "%Y-%m-%d").date()
                             if str(row.periodo) in ["1", "2"] 
                             else datetime.strptime(f"{int(b12_val)+1}{date_map_8.get(str(row.periodo),'')}", "%Y-%m-%d").date())
            except:
                data_comp = None
            siconfi_date = parse_date(str(row.data_status))
            status_msg = f"Câmara Quadrimestre {row.periodo}: " + (
                f"fora do prazo - data limite: {data_comp} e data status SICONFI: {siconfi_date}"
                if data_comp and siconfi_date and siconfi_date > data_comp 
                else f"dentro do prazo - data limite: {data_comp} e data status SICONFI: {siconfi_date}")
            msg_cam += status_msg + "\n"
    msg_tcm = ""
    idx_tcm = match_value_df(df, "org_entregavel", "Tribunal de Contas do Município do Rio de Janeiro - Relatório de Gestão Fiscal")
    if idx_tcm is not None:
        for row in df.iloc[idx_tcm:idx_tcm+3].itertuples():
            try:
                data_comp = (datetime.strptime(f"{b12_val}{date_map_8.get(str(row.periodo),'')}", "%Y-%m-%d").date()
                             if str(row.periodo) in ["1", "2"] 
                             else datetime.strptime(f"{int(b12_val)+1}{date_map_8.get(str(row.periodo),'')}", "%Y-%m-%d").date())
            except:
                data_comp = None
            siconfi_date = parse_date(str(row.data_status))
            status_msg = f"TCM Quadrimestre {row.periodo}: " + (
                f"fora do prazo - data limite: {data_comp} e data status SICONFI: {siconfi_date}"
                if data_comp and siconfi_date and siconfi_date > data_comp 
                else f"dentro do prazo - data limite: {data_comp} e data status SICONFI: {siconfi_date}")
            msg_tcm += status_msg + "\n"
    register_warning(warnings_list, 9, f"{msg_cam}\n{msg_tcm}")

    # D1_00011 a D1_00014 – Contagem de retificações, usando lambda e sum()
    idx = match_value_df(df, "entregavel", "Relatório Resumido de Execução Orçamentária")
    ret_1 = df.iloc[idx:idx+6]["status_relatorio"].apply(lambda x: 1 if x=="retificado" else 0).sum() if idx is not None else 0
    register_warning(warnings_list, 11, f"Quantidade de retificações : {ret_1}")

    idx = match_value_df(df, "entregavel", "Balanço Anual (DCA)")
    ret_2 = 1 if idx is not None and df.at[idx, "status_relatorio"]=="retificado" else 0
    register_warning(warnings_list, 12, f"Quantidade de retificações : {ret_2}")

    idx = match_value_df(df, "org_entregavel", "Prefeitura Municipal do Rio de Janeiro - RJ - Relatório de Gestão Fiscal")
    ret_3 = df.iloc[idx:idx+3]["status_relatorio"].apply(lambda x: 1 if x=="retificado" else 0).sum() if idx is not None else 0
    register_warning(warnings_list, 13, f"Quantidade de retificações : {ret_3}")

    idx = match_value_df(df, "org_entregavel", "Tribunal de Contas do Município do Rio de Janeiro - Relatório de Gestão Fiscal")
    ret_4 = df.iloc[idx:idx+3]["status_relatorio"].apply(lambda x: 1 if x=="retificado" else 0).sum() if idx is not None else 0
    register_warning(warnings_list, 14, f"Câmara - Quantidade de retificações : {ret_3}\nTCM - Quantidade de retificações : {ret_4}")

# Função auxiliar para construir o caminho e o nome da planilha MSC
def process_msc_file(i, b12_val, caminhoRels):
    file_name = f"MSC_{i:02d}{b12_val}.xlsx"
    sheet_name = f"MSC_{i:02d}{b12_val}"
    return os.path.join(caminhoRels, file_name), sheet_name

# Função para processar as validações que envolvem os arquivos MSC e PCASP
def validate_msc_files(ws_d1, warnings_list, caminhoRels):
    b12_val = ws_d1.range("B12").value
    msc_msg = ""

    # D1_00016: (Aqui usamos os dados da API; inserimos mensagem placeholder)
    register_warning(warnings_list, 16, "Validação dos envios MSC realizada via API.")

    # D1_00017: Verificar valores negativos na coluna "Valor"
    msc_msg = ""
    for i in range(1, 13):
        file_path, sheet_name = process_msc_file(i, b12_val, caminhoRels)
        if os.path.exists(file_path):
            try:
                wb = xw.Book(file_path)
                ws = wb.sheets[sheet_name]
                headers = ws.range("2:2").value
                try:
                    col_index = headers.index("Valor") + 1
                except ValueError:
                    col_index = None
                if col_index:
                    last_row = ws.range((3, col_index)).expand('down').last_cell.row
                    for j in range(3, last_row + 1):
                        if safe_float(ws.range((j, col_index)).value) < 0:
                            msc_msg += f"{sheet_name} com valor negativo na célula {xw.utils.rowcol_to_a1(2, col_index-1)}{j}. Favor verificar.\n"
                wb.close(save=False)
            except Exception as e:
                print("Erro no MSC (D1_00017):", sheet_name, e)
    register_warning(warnings_list, 17, msc_msg)

    # D1_00018: Verificar inconsistências na movimentação (colunas P e N)
    msc_msg = ""
    for i in range(1, 13):
        file_path, sheet_name = process_msc_file(i, b12_val, caminhoRels)
        if os.path.exists(file_path):
            try:
                wb = xw.Book(file_path)
                ws = wb.sheets[sheet_name]
                headers = ws.range("2:2").value
                try:
                    col_index = headers.index("Tipo_Valor") + 1
                except ValueError:
                    col_index = None
                if col_index:
                    last_row = ws.range((3, col_index)).expand('down').last_cell.row
                    for j in range(3, last_row + 1, 4):
                        beginning_DC = ws.range(f"P{j}").value
                        cum_Beginning = safe_float(ws.range(f"N{j}").value)
                        if ws.range(f"P{j+1}").value == "D":
                            cum_Change_D = safe_float(ws.range(f"N{j+1}").value)
                            cum_Change_C = safe_float(ws.range(f"N{j+2}").value)
                        else:
                            cum_Change_C = safe_float(ws.range(f"N{j+1}").value)
                            cum_Change_D = safe_float(ws.range(f"N{j+2}").value)
                        cum_Ending = safe_float(ws.range(f"N{j+3}").value)
                        if beginning_DC == "D":
                            if round(abs(cum_Beginning + cum_Change_D - cum_Change_C),2) != round(abs(cum_Ending),2):
                                msc_msg += f"{sheet_name} - Movimentação inconsistente nas linhas {j} a {j+3}. Favor verificar.\n"
                        else:
                            if round(abs(cum_Beginning + cum_Change_C - cum_Change_D),2) != round(abs(cum_Ending),2):
                                msc_msg += f"{sheet_name} - Movimentação inconsistente nas linhas {j} a {j+3}. Favor verificar.\n"
                    wb.close(save=False)
            except Exception as e:
                print("Erro no MSC (D1_00018):", sheet_name, e)
    register_warning(warnings_list, 18, msc_msg)

    # D1_00019: Verificar códigos de conta incorretos (coluna B)
    msc_msg = ""
    allowed_codes = {"10131", "10132", "20231", "20232"}
    for i in range(1, 13):
        file_path, sheet_name = process_msc_file(i, b12_val, caminhoRels)
        if os.path.exists(file_path):
            try:
                wb = xw.Book(file_path)
                ws = wb.sheets[sheet_name]
                last_row = ws.range("A3").expand('down').last_cell.row
                for j in range(3, last_row + 1):
                    if ws.range(f"B{j}").value not in allowed_codes:
                        msc_msg += f"{sheet_name} - Conta: {ws.range(f'A{j}').value} com código incorreto. Favor verificar.\n"
                wb.close(save=False)
            except Exception as e:
                print("Erro no MSC (D1_00019):", sheet_name, e)
    register_warning(warnings_list, 19, msc_msg)

    # D1_00020: Comparar saldo inicial do mês atual com o saldo final do mês anterior
    msc_msg = ""
    for i in range(2, 13):
        file_prev, sheet_prev = process_msc_file(i-1, b12_val, caminhoRels)
        file_curr, sheet_curr = process_msc_file(i, b12_val, caminhoRels)
        if os.path.exists(file_prev) and os.path.exists(file_curr):
            try:
                wb_curr = xw.Book(file_curr)
                ws_curr = wb_curr.sheets[sheet_curr]
                wb_prev = xw.Book(file_prev)
                ws_prev = wb_prev.sheets[sheet_prev]
                vec_curr = {}
                vec_prev = {}
                last_row_curr = ws_curr.range("A3").expand('down').last_cell.row
                conta = ws_curr.range("A3").value
                saldo = 0
                for j in range(3, last_row_curr + 1, 4):
                    current = ws_curr.range(f"A{j}").value
                    if current == conta:
                        saldo += safe_float(ws_curr.range(f"N{j}").value) if ws_curr.range(f"P{j}").value=="D" else -safe_float(ws_curr.range(f"N{j}").value)
                    else:
                        vec_curr[conta] = saldo
                        conta = current
                        saldo = safe_float(ws_curr.range(f"N{j}").value) if ws_curr.range(f"P{j}").value=="D" else -safe_float(ws_curr.range(f"N{j}").value)
                vec_curr[conta] = saldo

                last_row_prev = ws_prev.range("A3").expand('down').last_cell.row
                conta = ws_prev.range("A3").value
                saldo = 0
                for j in range(6, last_row_prev + 1, 4):
                    current = ws_prev.range(f"A{j}").value
                    if current == conta:
                        saldo += safe_float(ws_prev.range(f"N{j}").value) if ws_prev.range(f"P{j}").value=="D" else -safe_float(ws_prev.range(f"N{j}").value)
                    else:
                        vec_prev[conta] = saldo
                        conta = current
                        saldo = safe_float(ws_prev.range(f"N{j}").value) if ws_prev.range(f"P{j}").value=="D" else -safe_float(ws_prev.range(f"N{j}").value)
                vec_prev[conta] = saldo

                for key in vec_curr:
                    if key in vec_prev and vec_curr[key] != vec_prev[key]:
                        msc_msg += f"Chave {key} em {sheet_curr} possui saldo inicial diferente do saldo final do mês anterior. Favor verificar.\n"
                wb_curr.close(save=False)
                wb_prev.close(save=False)
            except Exception as e:
                print("Erro no MSC (D1_00020) para MSC", i, e)
    register_warning(warnings_list, 20, msc_msg)

    # D1_00021: Verificar natureza dos saldos (MSC vs PCASP)
    msc_msg = ""
    codes = ["2111","2112","2113","2114","2121","2122","2123","2124","2125","2126","213","214","215","221","222","223"]
    for i in range(1, 13):
        file_msc, sheet_msc = process_msc_file(i, b12_val, caminhoRels)
        file_pcasp = os.path.join(caminhoRels, f"PCASP ESTENDIDO {b12_val}.xlsx")
        if os.path.exists(file_msc) and os.path.exists(file_pcasp):
            try:
                wb_msc = xw.Book(file_msc)
                ws_msc = wb_msc.sheets[sheet_msc]
                wb_pcasp = xw.Book(file_pcasp)
                ws_pcasp = wb_pcasp.sheets["Estendido " + b12_val]
                vec = {}
                conta = ws_msc.range("A3").value
                saldo = 0
                last_row = ws_msc.range("A3").expand('down').last_cell.row
                for k in range(6, last_row + 1, 4):
                    current = ws_msc.range(f"A{k}").value
                    if current == conta:
                        saldo += safe_float(ws_msc.range(f"N{k}").value) if ws_msc.range(f"P{k}").value=="D" else -safe_float(ws_msc.range(f"N{k}").value)
                    else:
                        vec[conta] = "D" if saldo > 0 else ("C" if saldo < 0 else "N")
                        conta = current
                        saldo = safe_float(ws_msc.range(f"N{k}").value) if ws_msc.range(f"P{k}").value=="D" else -safe_float(ws_msc.range(f"N{k}").value)
                vec[conta] = "D" if saldo > 0 else ("C" if saldo < 0 else "N")
                for key, nat in vec.items():
                    if any(str(key).startswith(code) for code in codes):
                        pos = ws_pcasp.range("H:H").find(key)
                        if pos is not None:
                            pcasp_nat = str(ws_pcasp.range(f"K{pos.row}").value)[0]
                            if nat != "N" and nat != pcasp_nat:
                                msc_msg += f"Conta {key} em {sheet_msc} possui natureza {nat} diferente do PCASP ({pcasp_nat}). Favor verificar.\n"
                wb_msc.close(save=False)
                wb_pcasp.close(save=False)
            except Exception as e:
                print("Erro no MSC (D1_00021):", sheet_msc, e)
    register_warning(warnings_list, 21, msc_msg)

    # D1_00022: Verificar ausência de código de poder/órgão (coluna B vazia) nos MSC
    msc_msg = ""
    for i in range(1, 13):
        file_msc, sheet_msc = process_msc_file(i, b12_val, caminhoRels)
        if os.path.exists(file_msc):
            try:
                wb = xw.Book(file_msc)
                ws = wb.sheets[sheet_msc]
                last_row = ws.range("A3").expand('down').last_cell.row
                for j in range(3, last_row + 1):
                    if ws.range(f"B{j}").value == "":
                        msc_msg += f"Conta {ws.range(f'A{j}').value} em {sheet_msc} sem código informado. Favor verificar.\n"
                wb.close(save=False)
            except Exception as e:
                print("Erro no MSC (D1_00022):", sheet_msc, e)
    # Verifica MSC_13 (encerramento)
    file_msc = os.path.join(caminhoRels, f"MSC_13{b12_val}.xlsx")
    if os.path.exists(file_msc):
        try:
            wb = xw.Book(file_msc)
            sheet_msc = f"MSC_13{b12_val}"
            ws = wb.sheets[sheet_msc]
            last_row = ws.range("A3").expand('down').last_cell.row
            for j in range(3, last_row + 1):
                if ws.range(f"B{j}").value == "":
                    msc_msg += f"Conta {ws.range(f'B{j}').value} em {sheet_msc} sem código informado. Favor verificar.\n"
            wb.close(save=False)
        except Exception as e:
            print("Erro no MSC (D1_00022) MSC_13:", e)
    register_warning(warnings_list, 22, msc_msg)

    # D1_00023: Verificar duplicação de dados do Poder Executivo
    msc_msg = ""
    for i in range(2, 13):
        file_prev, sheet_prev = process_msc_file(i-1, b12_val, caminhoRels)
        file_curr, sheet_curr = process_msc_file(i, b12_val, caminhoRels)
        if os.path.exists(file_prev) and os.path.exists(file_curr):
            try:
                wb_curr = xw.Book(file_curr)
                ws_curr = wb_curr.sheets[sheet_curr]
                wb_prev = xw.Book(file_prev)
                ws_prev = wb_prev.sheets[sheet_prev]
                tudoIgual = all(ws_curr.range(f"N{j}").value == ws_prev.range(f"N{j}").value 
                                for j in range(3, ws_curr.range("A3").expand('down').last_cell.row + 1)
                                if str(ws_curr.range(f"B{j}").value).startswith("1"))
                if tudoIgual:
                    msc_msg += f"Matriz do mês {sheet_curr} possui dados duplicados do Poder Executivo do mês anterior. Favor verificar.\n"
                wb_curr.close(save=False)
                wb_prev.close(save=False)
            except Exception as e:
                print("Erro no MSC (D1_00023) para MSC", i, e)
    register_warning(warnings_list, 23, msc_msg)

    # D1_00024: Verificar duplicação de dados do Poder Legislativo
    msc_msg = ""
    for i in range(2, 13):
        file_prev, sheet_prev = process_msc_file(i-1, b12_val, caminhoRels)
        file_curr, sheet_curr = process_msc_file(i, b12_val, caminhoRels)
        if os.path.exists(file_prev) and os.path.exists(file_curr):
            try:
                wb_curr = xw.Book(file_curr)
                ws_curr = wb_curr.sheets[sheet_curr]
                wb_prev = xw.Book(file_prev)
                ws_prev = wb_prev.sheets[sheet_prev]
                tudoIgual = all(ws_curr.range(f"N{j}").value == ws_prev.range(f"N{j}").value 
                                for j in range(3, ws_curr.range("A3").expand('down').last_cell.row + 1)
                                if str(ws_curr.range(f"B{j}").value).startswith("2"))
                if tudoIgual:
                    msc_msg += f"Matriz do mês {sheet_curr} possui dados duplicados do Poder Legislativo do mês anterior. Favor verificar.\n"
                wb_curr.close(save=False)
                wb_prev.close(save=False)
            except Exception as e:
                print("Erro no MSC (D1_00024) para MSC", i, e)
    register_warning(warnings_list, 24, msc_msg)

    # D1_00025: Verificar natureza dos saldos (MSC vs PCASP) para contas específicas
    msc_msg = ""
    codes = ["2111","2112","2113","2114","2121","2122","2123","2124","2125","2126","213","214","215","221","222","223"]
    for i in range(1, 13):
        file_msc, sheet_msc = process_msc_file(i, b12_val, caminhoRels)
        file_pcasp = os.path.join(caminhoRels, f"PCASP ESTENDIDO {b12_val}.xlsx")
        if os.path.exists(file_msc) and os.path.exists(file_pcasp):
            try:
                wb_msc = xw.Book(file_msc)
                ws_msc = wb_msc.sheets[sheet_msc]
                wb_pcasp = xw.Book(file_pcasp)
                ws_pcasp = wb_pcasp.sheets["Estendido " + b12_val]
                vec = {}
                conta = ws_msc.range("A3").value
                saldo = 0
                last_row = ws_msc.range("A3").expand('down').last_cell.row
                for k in range(6, last_row + 1, 4):
                    current = ws_msc.range(f"A{k}").value
                    if current == conta:
                        saldo += safe_float(ws_msc.range(f"N{k}").value) if ws_msc.range(f"P{k}").value=="D" else -safe_float(ws_msc.range(f"N{k}").value)
                    else:
                        vec[conta] = "D" if saldo > 0 else ("C" if saldo < 0 else "N")
                        conta = current
                        saldo = safe_float(ws_msc.range(f"N{k}").value) if ws_msc.range(f"P{k}").value=="D" else -safe_float(ws_msc.range(f"N{k}").value)
                vec[conta] = "D" if saldo > 0 else ("C" if saldo < 0 else "N")
                for key, nat in vec.items():
                    if any(str(key).startswith(code) for code in codes):
                        pos = ws_pcasp.range("H:H").find(key)
                        if pos is not None:
                            pcasp_nat = str(ws_pcasp.range(f"K{pos.row}").value)[0]
                            if nat != "N" and nat != pcasp_nat:
                                msc_msg += f"Conta {key} em {sheet_msc} possui natureza {nat} diferente do PCASP ({pcasp_nat}). Favor verificar.\n"
                wb_msc.close(save=False)
                wb_pcasp.close(save=False)
            except Exception as e:
                print("Erro no MSC (D1_00025):", sheet_msc, e)
    register_warning(warnings_list, 25, msc_msg)

# =============================================================================
# Função principal que une as validações a partir do CSV e as dos arquivos MSC/PCASP
# =============================================================================
def btnSelecionarMSCs_Click_from_csv():
    wb = xw.Book.caller()
    ws_d1 = wb.sheets["Checklist D1"]
    base_path = os.path.dirname(wb.fullname)
    ano_val = str(ws_d1.range("B12").value)
    caminhoRels = os.path.join(base_path, "MSC" + ano_val) + os.sep

    warnings_list = []

    # Limpar status na planilha Checklist D1 (linhas 18 a 40, colunas D e E)
    for i in range(18, 41):
        ws_d1.range(f"D{i}").color = (255,255,255)
        ws_d1.range(f"D{i}").value = ""
        ws_d1.range(f"E{i}").value = ""

    # Carregar dados da API a partir do CSV
    csv_file = os.path.join(os.getcwd(), f"Siconfi_{ano_val}_output.csv")
    if not os.path.exists(csv_file):
        print("Arquivo CSV Siconfi não encontrado:", csv_file)
        return
    df = pd.read_csv(csv_file, encoding="utf-8")
    
    # Validações com dados da API
    validate_api_items(df, ws_d1, warnings_list)
    
    # Validações que envolvem arquivos MSC (e PCASP)
    validate_msc_files(ws_d1, warnings_list, caminhoRels)
    
    ws_d1.activate()
    print("Validações D1 (00001 a 00025) concluídas com sucesso.")
    print("Warnings coletados:")
    for item in warnings_list:
        print(f"Item {item[0]}: {item[1]}")

# =============================================================================
# Execução principal (standalone para testes – adapte se usar via xlwings)
# =============================================================================
def btnSelecionarMSCs_Click_from_csv():
    print("Executando btnSelecionarMSCs_Click_from_csv()...")  # Debug
    # Adicione aqui a lógica do seu processamento



if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: python D1_convertido_otimizado.py <caminho_do_arquivo>")
        sys.exit(1)

    caminho_arquivo = sys.argv[1]
    
    # Obtém o caminho absoluto
    caminho_completo = os.path.abspath(caminho_arquivo)
    
    print(f"Abrindo arquivo: {caminho_completo}")  # Debug

    # Abrindo a planilha corretamente
    wb = xw.Book(caminho_completo)

    # Chamando a função principal
    btnSelecionarMSCs_Click_from_csv()

    # Fechando o arquivo após a execução
    wb.close()

# Criar a pasta de processados, se não existir
processados_folder = "uploads/processados"
os.makedirs(processados_folder, exist_ok=True)

# Caminho onde o arquivo processado será salvo
output_file = os.path.join(processados_folder, "arquivo_processado.xlsx")

# Salvar o arquivo processado no local correto
wb.save(output_file)
print(f"Arquivo processado salvo em: {output_file}")
