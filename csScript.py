from demoparser2 import DemoParser
from tkinter import Tk
from tkinter.filedialog import askopenfilenames, asksaveasfilename
import pandas as pd
import os

def main():
    print("Instruções: Selecione todas as demos e clique em abrir")
    print("OBS: Antes de começar, renomeie todas as demos para o nome do mapa, isso será importante no resultado final. Caso já tenha feito isso, aperte ENTER")
    input()

    root = Tk()
    root.withdraw()

    demo_paths = askopenfilenames(filetypes=[("Demo files", "*.dem")])
    if not demo_paths:
        print("Nenhum arquivo selecionado.")
        return
    else:
        print("Escolha o local para salvar e o nome, após escolher, aguarde")

    # Caminho para salvar o Excel
    output_path = asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")],
        title="Salvar como",
        initialfile="TabelaDemos.xlsx"
    )
    if not output_path:
        print("Nenhum local de salvamento selecionado.")
        return

    all_players = set()
    for demo_path in demo_paths:
        parser = DemoParser(demo_path)
        max_tick = parser.parse_event("round_end")["tick"].max()
        df_stats = parser.parse_ticks(["name"], ticks=[max_tick])
        for name in df_stats["name"]:
            all_players.add(name)

    all_players = sorted(all_players)
    print("Jogadores encontrados nas demos:")
    for i, p in enumerate(all_players, start=1):
        print(f"{i}: {p}")

    time1_input = input("\nDigite, SEM ESPAÇO, os números dos jogadores do TIME 1 separados por vírgula (ex: 1,3,5):\n")
    indices = [int(x.strip()) - 1 for x in time1_input.split(",")]
    time1_players = {all_players[i] for i in indices if 0 <= i < len(all_players)}
    time2_players = set(all_players) - time1_players

    print(f"\nTIME 1: {sorted(time1_players)}")
    print(f"TIME 2: {sorted(time2_players)}")

    writer = pd.ExcelWriter(output_path, engine='openpyxl')
    row_start = 0
    all_time1_stats = []
    all_time2_stats = []

    for demo_path in demo_paths:
        print(f"Processando: {demo_path}")
        parser = DemoParser(demo_path)
        max_tick = parser.parse_event("round_end")["tick"].max()

        wanted_fields = [
            "kills_total",
            "deaths_total",
            "assists_total",
            "mvps",
            "headshot_kills_total",
            "ace_rounds_total",
            "4k_rounds_total",
            "3k_rounds_total",
            "name"
        ]

        df_stats = parser.parse_ticks(wanted_fields, ticks=[max_tick])

        if 'tick' in df_stats.columns:
            df_stats = df_stats.drop(columns=['tick'])
        if 'steamid' in df_stats.columns:
            df_stats = df_stats.drop(columns=['steamid'])

        rename_map = {
            "name": "Nome",
            "kills_total": "Kills",
            "deaths_total": "Deaths",
            "assists_total": "Assists",
            "headshot_kills_total": "Kills HS",
            "ace_rounds_total": "5K",
            "4k_rounds_total": "4K",
            "3k_rounds_total": "3K",
            "mvps": "MVPS"
        }
        df_stats = df_stats.rename(columns=rename_map)

        cols = df_stats.columns.tolist()
        if "Nome" in cols:
            cols.insert(0, cols.pop(cols.index("Nome")))
        df_stats = df_stats[cols]

        df_time1 = df_stats[df_stats["Nome"].isin(time1_players)].reset_index(drop=True)
        df_time2 = df_stats[df_stats["Nome"].isin(time2_players)].reset_index(drop=True)

        df_time1["KDA"] = (df_time1["Kills"] + df_time1["Assists"]) / df_time1["Deaths"].replace(0, 1)
        df_time2["KDA"] = (df_time2["Kills"] + df_time2["Assists"]) / df_time2["Deaths"].replace(0, 1)

        df_time1["KDA"] = df_time1["KDA"].round(2)
        df_time2["KDA"] = df_time2["KDA"].round(2)

        demo_name = os.path.splitext(os.path.basename(demo_path))[0]
        sheet_name = 'Estatísticas'

        header_df = pd.DataFrame([[demo_name]], columns=[df_stats.columns[0]])
        header_df.to_excel(writer, sheet_name=sheet_name, startrow=row_start, index=False, header=False)

        if not df_time1.empty:
            df_time1.to_excel(writer, sheet_name=sheet_name, startrow=row_start + 1, index=False)
            row_start += len(df_time1) + 2
            all_time1_stats.append(df_time1)
        else:
            row_start += 2

        if not df_time2.empty:
            df_time2.to_excel(writer, sheet_name=sheet_name, startrow=row_start, index=False)
            row_start += len(df_time2) + 3
            all_time2_stats.append(df_time2)
        else:
            row_start += 3

    if all_time1_stats:
        total_time1 = pd.concat(all_time1_stats).groupby("Nome", as_index=False).sum(numeric_only=True)
        total_time1["KDA"] = (total_time1["Kills"] + total_time1["Assists"]) / total_time1["Deaths"].replace(0, 1)
        total_time1["KDA"] = total_time1["KDA"].round(2)

    if all_time2_stats:
        total_time2 = pd.concat(all_time2_stats).groupby("Nome", as_index=False).sum(numeric_only=True)
        total_time2["KDA"] = (total_time2["Kills"] + total_time2["Assists"]) / total_time2["Deaths"].replace(0, 1)
        total_time2["KDA"] = total_time2["KDA"].round(2)

    row_start += 2
    pd.DataFrame([["TOTAL TIME 1"]], columns=["Nome"]).to_excel(writer, sheet_name=sheet_name, startrow=row_start, index=False, header=False)
    row_start += 1
    total_time1.to_excel(writer, sheet_name=sheet_name, startrow=row_start, index=False)
    row_start += len(total_time1) + 2

    pd.DataFrame([["TOTAL TIME 2"]], columns=["Nome"]).to_excel(writer, sheet_name=sheet_name, startrow=row_start, index=False, header=False)
    row_start += 1
    total_time2.to_excel(writer, sheet_name=sheet_name, startrow=row_start, index=False)

    writer.close()
    print(f"\nPlanilha criada com sucesso: {output_path}")
    input("\nPressione Enter para sair...")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"Erro: {e}")
        input("\nPressione Enter para sair...")
