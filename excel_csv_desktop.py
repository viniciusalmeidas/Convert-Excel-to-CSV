
from tkinter import messagebox
import pandas as pd
import os 
from tkinter import * 
from tkinter import filedialog
import glob

def convert():
    error = 0
    tamanho= len(os.listdir(origemValue.get().replace(os.sep,'/')))
    for filename in os.listdir(origemValue.get().replace(os.sep,'/')):    
        file = os.path.join(origemValue.get(),filename).replace(os.sep,'/')
        if os.path.isfile(file):
            print(file)
            destino = destinoValue.get().replace(os.sep,'/')
            try:
                read_file = pd.read_excel(file)
                read_file.to_csv(f'{destino}\{os.path.splitext(filename)[0]}.csv', index=None, header=True)    
            except:
                error = error + 1
                messagebox.showwarning("Error",f'Erro do arquivo: {filename}')
    
    
    if joinVar.get() ==1:
        try:
            combineCsv(destino)
        except:
            messagebox.showerror("Error",f'Erro ao juntar os arquivos em  {destino}')
    
    Label(fg="green",text=f'Conversão Concluída com sucesso. {tamanho-error} arquivo(s) convertido(s) e {error} erro(s)', font="arial 10").place(x=50,y=150)
    OrigemEntry.delete(0, 'end')
    DestinoEntry.delete(0, 'end')

def askFolderOrig():
    if 1 == 1:
        origemValue = filedialog.askdirectory()
        OrigemEntry.insert(END,""+origemValue)
    else: 
        origemValue = filedialog.askdirectory()
        OrigemEntry.insert(END,""+origemValue)

def askFolderDest():
    if 1 == 1:
        destinoValue = filedialog.askdirectory()
        DestinoEntry.insert(END,""+destinoValue)
    else: 
        destinoValue = filedialog.askdirectory()
        DestinoEntry.insert(END,""+destinoValue)

def combineCsv(directory):
    extension = 'csv'
    print(directory)
    os.chdir(directory)
    all_filenames = [i for i in glob.glob('*.{}'.format(extension))]
    #combine all files in the list
    combined_csv = pd.concat([pd.read_csv(f) for f in all_filenames ])
    #export to csv
    combined_csv.to_csv( "ACOPLADO_csv.csv", index=False, encoding='utf-8-sig')

def main():
    global root
    root=Tk()
    root.title("Excel to CSV - @viniciusalmeidas v.almsou@uol.com.br")
    root.geometry("525x230")
    root.resizable(False, False)
    # root.iconbitmap('icone.ico')

    #Cabeçalho
    Label(root, text="Converter Excel para .csv", font="arial 13").pack(pady=25)

    #Linha 1
    global OrigemEntry
    global origemValue
    Label(text="Pasta de origem:", font="arial 10").place(x=20,y=80)
    origemValue=StringVar()
    OrigemEntry=Entry(root,textvariable=origemValue,width=35,bd=2,font="arial 9")
    OrigemEntry.place(x=160,y=80)
    Button(text="Selecione...",font="arial 9",width=10,height=0,command=askFolderOrig).place(x=425,y=75)

    #Linha 2
    global destinoValue
    global DestinoEntry
    Label(text="Pasta de destino:", font="arial 10").place(x=20,y=110)
    destinoValue=StringVar()
    DestinoEntry=Entry(root,textvariable=destinoValue,width=35,bd=2,font="arial 9")
    DestinoEntry.place(x=160,y=120)
    Button(text="Selecione...",font="arial 9",width=10,height=0,command=askFolderDest).place(x=425,y=115)

    global joinVar
    joinVar = IntVar()
    Checkbutton(root, text='Juntar arquivos *.csv ao finalizar',variable=joinVar, onvalue=1, offvalue=0).place(x=40,y=175)

    # Submit Button
    Button(text="Converter",font=3,width=8,height=1,command=convert).place(x=390,y=170)
    # Exit Button
    Button(text="Sair",font=3,width=8,height=1, command=root.destroy).place(x=270,y=170)

    root.mainloop()
    

if __name__ =='__main__':
    main()
