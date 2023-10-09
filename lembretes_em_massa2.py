import win32com.client;
from win32com.client import Dispatch;
import pandas as pd;
import datetime;
import tkinter as tk;
from tkinter import *;
from tkcalendar import DateEntry;

window = tk.Tk();
window.title("Lembretes em massa");
window.geometry("400x200+683+384");
window.iconbitmap("lembrete.ico");
style = ttk.Style();
style.theme_use('clam');
style.configure('my.DateEntry',
                fieldbackground='white',
                background='light green',
                foreground='black',
                arrowcolor='grey');
lblCalInicio = tk.Label(text="Data de início:")
lblCalInicio.pack()
calendarioSelecaoInicio = DateEntry(style='my.DateEntry');
horaInicio = "";
minutoInicio = "";
segundoInicio = "";
sbHora = Spinbox(window,
    from_=0,
    to=23,
    wrap=True,
    textvariable=horaInicio,
    width=2,
    state="readonly",
    font=f,
    justify=tk.CENTER);
calendarioSelecaoInicio.pack();
lblCalFim = tk.Label(text="Data de término:")
lblCalFim.pack();
calendarioSelecaoFim = DateEntry(style='my.DateEntry');
calendarioSelecaoFim.pack();
lblCategoria = tk.Label(text="Categoria do lembrete:")
lblCategoria.pack();
entradaCategoria = tk.Entry();
entradaCategoria.pack();
categoriaLembrete = "";
dataInicio = "";
dataFim = "";
def obterData():
    outlook = win32com.client.Dispatch("Outlook.Application");
    espacoMestre = outlook.GetNamespace("MAPI");
    dataSelecaoInicio = calendarioSelecaoInicio.get_date();
    dataSelecaoFim = calendarioSelecaoFim.get_date();
    categoriaLembrete = entradaCategoria.get();
    categoriaExiste = False;
    for cat in espacoMestre.session.categories:
        if (cat == categoriaLembrete):
            categoriaExiste = True;
    if (categoriaExiste == False):
        tk.messagebox.showerror(title="Categoria não existente", message="A categoria informada não existe, favor informar novamente.");
        return;
    dataInicio = datetime.datetime(dataSelecaoInicio.year, dataSelecaoInicio.month, dataSelecaoInicio.day);
    dataFim = datetime.datetime(dataSelecaoFim.year, dataSelecaoFim.month, dataSelecaoFim.day);

    dias_Uteis = pd.bdate_range(dataInicio, dataFim);
    num_meses = (dataFim.year-dataInicio.year) * 12 + (dataFim.month-dataInicio.month);

    meses_Registro = [];
    datas_Lembretes = [];
    cont = 0;

    #Iteração para adicionar na lista meses_registro os números dos meses
    #dentro do intervalo de tempo fornecido.
    while cont <= (num_meses):
        if (cont == 0):
            meses_Registro.append(dataInicio.month);
        else:
            meses_Registro.append(dataInicio.month + cont);
        cont = cont + 1;


    #Itera por todos os meses dentro do intervalo das datas informas e itera também
    #por todos os dias úteis dentro do mesmo intervalo. Dentro de cada iteração dos dias úteis
    #compara os dias armazenados na variavel ultimoDiaUtil com o dia da iteração atual e verifica
    #se é maior, se for atribui na variável, se não, pula o dia.
    #Ao final da iteração de cada mês, é atribuído na lista datas_Lembretes os primeiros e últimos dias úteis
    #de cada mês.
    cont = 0;
    horaInicioLembrete = pd.Timedelta(hours=8)
    for mes in meses_Registro:
        for dia in dias_Uteis:
            if (dia.month == mes and cont == 0):
                primeiroDiaUtil = dia;
                ultimoDiaUtil = dia;
                cont = cont + 1;
            elif (dia.month == mes and dia.day > ultimoDiaUtil.day):
                ultimoDiaUtil = dia;
        if (meses_Registro[0] == mes and dataInicio.day > primeiroDiaUtil.day):
            datas_Lembretes.append(ultimoDiaUtil+horaInicioLembrete);
        elif (meses_Registro[len(meses_Registro)-1] == mes and dataFim.day < ultimoDiaUtil.day):
            datas_Lembretes.append(primeiroDiaUtil+horaInicioLembrete);
        else:
            datas_Lembretes.append(primeiroDiaUtil+horaInicioLembrete);
            datas_Lembretes.append(ultimoDiaUtil+horaInicioLembrete);
        cont = 0;


    #Cria os lembretes no calendário do Outlook

    for data in datas_Lembretes:
        appt = outlook.CreateItem(1);
        appt.Start = data.strftime("%Y-%m-%d %X") # yyyy-MM-dd hh:mm;
        if (data.day < 15):
            appt.Subject = "Fechar período do PAC";
        elif (data.day > 15):
            appt.Subject = "Abrir período do PAC";
        appt.Categories = categoriaLembrete;
        appt.Duration = 600;
        appt.Save();
        appt.Send();

btnOk = tk.Button(window, text="OK", command=obterData);
btnOk.pack(pady=10);

window.mainloop();





