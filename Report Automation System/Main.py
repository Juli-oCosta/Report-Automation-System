import os
import pandas as pd
import win32com.client as win32

data = pd.read_excel('PlanilhasExcel/dados.xlsx')

data["Início do atendimento"] = pd.to_datetime(data["Início do atendimento"], format="%H:%M:%S", errors='coerce')
data["Final do atendimento"] = pd.to_datetime(data["Final do atendimento"], format="%H:%M:%S", errors='coerce')

data["Final do atendimento"] = data.apply(
     lambda row: row["Final do atendimento"] + pd.Timedelta(days=1) if row["Final do atendimento"] < row["Início do atendimento"] else row["Final do atendimento"],
     axis=1
)

data["Tempo"] = data["Final do atendimento"] - data["Início do atendimento"]
tempo_medio = data.groupby("Atendente")["Tempo"].mean()

quantidade_demandas = data.groupby("Atendente").size()
data["Demandas"] = data.groupby("Atendente")["ID da demanda"].transform("size")

demandas_abertas = data[data["Final do atendimento"].isna()].groupby("Atendente").size()
data["Demandas abertas"] = data["Atendente"].map(demandas_abertas).fillna(0).astype(int)

print("Quantidade de demandas por atendente:")
print(quantidade_demandas)

print("\nTempo médio por atendente:")
print(tempo_medio)

print("\nDemandas em aberto por atendente:")
print(demandas_abertas)

relatorios_por_gestor = data.groupby("e-mail gestor")

for email, relatorio in relatorios_por_gestor:
    nome_arquivo = f"DadosCSV/{email}_relatorio.csv"
    relatorio.to_csv(nome_arquivo, index=False)

for email, relatorio in relatorios_por_gestor:

    outlook = win32.Dispatch("outlook.application")
    email_outlook = outlook.CreateItem(0)

    email_outlook.To = email
    email_outlook.Subject = "Relatório Automático"
    email_outlook.HTMLBody = f"""
    <p>Olá,</p>
    <p>Segue o relatório atualizado das demandas.</p>
    <p>Atenciosamente,</p>
    <p>(Responsável pelo sistema de relatórios).</p>
    """

    nome_arquivo_absoluto = os.path.abspath(nome_arquivo)

    email_outlook.Attachments.Add(nome_arquivo_absoluto)

    email_outlook.Send()
    print(f"E-mail enviado para {email} com o relatório {nome_arquivo}")