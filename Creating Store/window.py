from tkinter import *
import pandas as pd
from tkinter import filedialog
from tkinter.filedialog import askopenfilename

def generateStoreFile(sName,gui):
    print("Generate Store File"+sName)
    selected_text_list = [listbox.get(i) for i in listbox.curselection()]
    articles=[]
    sku=[]
    qty=[]
    for each in selected_text_list:
        val=mainFileData[mainFileData[mainFileData.columns[1]]==each]
        index=val.index[0]
        x=mainFileData.iloc[[index]]
        articles.append(x[x.columns[1]].values[0])
        sku.append(str(sName)+x[x.columns[2]].values[0])
        qty.append(x[x.columns[3]].values[0])
        
    df=pd.DataFrame({
        mainFileData.columns[1]:articles,
        mainFileData.columns[2]:sku,
        mainFileData.columns[3]:qty,
        })    
    path=filedialog.askdirectory()
    if(len(path)!=0):  
        writer = pd.ExcelWriter(path+"\\"+str(sName)+'.xlsx', engine='xlsxwriter')
        df.to_excel(writer, sheet_name='Sheet1', index=False)
        writer.save()
        gui.withdraw()
    print(path)
    # path=r"C:\Users\Furqan\Desktop\Chachu Ramzan\New Sales Software\newStores"
    
    # folder_selected = root.askdirectory()
    # writer, sheet_name='Sheet1'
    # writer.save()
    print(df)
    # for each in newList:
    #     print(each[each.columns[1]])
        
            

def newStoreLayout(sName):
    gui = Toplevel()

    gui.geometry("1000x600")
    gui.title("FT")
    
    
    canvas = Canvas(
    gui,
    bg = "#ffffff",
    height = 600,
    width = 1000,
    bd = 0,
    highlightthickness = 0,
    relief = "ridge")
    canvas.place(x = 0, y = 0)
    
    
    # #choose articles text
    background_img = PhotoImage(file = f"background01.png")
    background01 = canvas.create_image(
        120,30,
        image=background_img)
    
    # # generate store file Button
    img2 = PhotoImage(file = f"img2.png")
    b2 = Button(
        gui,
        image = img2,
        borderwidth = 0,
        highlightthickness = 0,
        command =lambda: generateStoreFile(sName,gui),
        relief = "flat")
    b2.place(
    x = 650, y = 10,
    width = 307,
    height = 46)
    
    #choosing articles list
    global listbox
    listbox=Listbox(gui,bg="#ffffff",selectmode="multiple",font=('Times',25))
    listbox.place(x=10,y=60,width=960,height=540)
    scrollbar=Scrollbar(gui)
    scrollbar.pack(side=RIGHT,fill=BOTH)
    
    
    for values in mainFileData[mainFileData.columns[1]]: 
        listbox.insert(END,values)
    
    # x=60    
    # for values in mainFileData[mainFileData.columns[1]]: 
    #     print(values)
    #     listbox.insert(END,Checkbutton(gui, text = values).place(x=10,y=0+x))
    #     x=x+50
    
        
        
    listbox.config(yscrollcommand = scrollbar.set)
    scrollbar.config(command = listbox.yview)
    
    
    
    gui.mainloop()




def openMainFile():
    print("Opening File")
    global filename,mainFileData
    filename = askopenfilename()
    # filename=r"C:\Users\Furqan\Desktop\Chachu Ramzan\New Sales Software\FIRST-FILE-TO-UPLOAD.xlsx"
    mainFileData=pd.read_excel(filename)
    # for values in mainFileData[mainFileData.columns[1]]: 
    #     print(values)
    

def creatStore():
    global newStoreName
    newStoreName=''
    if entry0.get() != '':
        newStoreName=entry0.get()
        print(newStoreName)
        newStoreLayout(newStoreName)
        entry0.delete(0,END)







window = Tk()

window.geometry("1000x600")
window.configure(bg = "#ffffff")
window.title("FT")
canvas = Canvas(
    window,
    bg = "#ffffff",
    height = 600,
    width = 1000,
    bd = 0,
    highlightthickness = 0,
    relief = "ridge")
canvas.place(x = 0, y = 0)

img0 = PhotoImage(file = f"img0.png")
b0 = Button(
    image = img0,
    borderwidth = 0,
    highlightthickness = 0,
    command = openMainFile,
    relief = "flat")

b0.place(
    x = 347, y = 269,
    width = 307,
    height = 46)

img1 = PhotoImage(file = f"img1.png")
b1 = Button(
    image = img1,
    borderwidth = 0,
    highlightthickness = 0,
    command = creatStore,
    relief = "flat")

b1.place(
    x = 347, y = 431,
    width = 307,
    height = 46)

entry0_img = PhotoImage(file = f"img_textBox0.png")
entry0_bg = canvas.create_image(
    500.5, 392.0,
    image = entry0_img)

entry0 = Entry(
    bd = 0,
    bg = "#dcdcdc",
    highlightthickness = 0)

entry0.place(
    x = 355.0, y = 369,
    width = 291.0,
    height = 44)

background_img = PhotoImage(file = f"background.png")
background = canvas.create_image(
    500.0, 216.0,
    image=background_img)

window.resizable(False, False)
window.mainloop()
