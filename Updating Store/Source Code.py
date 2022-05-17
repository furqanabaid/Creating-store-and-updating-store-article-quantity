from tkinter import *
import pandas as pd
from tkinter import filedialog
from tkinter.filedialog import askopenfilename
import glob
import openpyxl as excel
import os
from tkinter import messagebox
from pathlib import Path

def openMainFile():
    print("Opening File")
    global filename,mainFileData,skuAndQtyDict
    filename = askopenfilename()
    # filename=r"C:\Users\Furqan\Desktop\Chachu Ramzan\New Sales Software\FIRST-FILE-TO-UPLOAD.xlsx"
    mainFileData=pd.read_excel(filename)
    sku=[i for i in mainFileData[mainFileData.columns[2]]]
    qty=[i for i in mainFileData[mainFileData.columns[3]]]
    skuAndQtyDict={mainFileData.columns[2]:sku,mainFileData.columns[3]:qty}
    # print(skuAndQtyDict[mainFileData.columns[2]][0],skuAndQtyDict[mainFileData.columns[3]][0])
    

def ChooseStoreLocation():
    global storesPath
    print("location getting")
    storesPath=filedialog.askdirectory()
    print(storesPath)

def updateAllStores():
    print("Updating all stores")
    updatedStoresPath=filedialog.askdirectory()
    print(updatedStoresPath)
    extension='xlsx'
    os.chdir(storesPath)
    files=[i for i in glob.glob('*.{}'.format(extension))]
    for f in files:
        boo=False
        nfName=storesPath+'/'+f
        unfName=updatedStoresPath+'/'+f
        path=Path(nfName)
        path1=Path(unfName)
        store=pd.read_excel(path)
        articles=[i for i in store[store.columns[0]]]
        sku=[i for i in store[store.columns[1]]]
        qty=[i for i in store[store.columns[2]]]
        storeSkuAndQtyDict={store.columns[0]:articles,store.columns[1]:sku,store.columns[2]:qty}
        for idx,l in enumerate(store[store.columns[1]]):
            for indx,saqd in enumerate(skuAndQtyDict[mainFileData.columns[2]]) : 
                if saqd in l:
                    print(storeSkuAndQtyDict[store.columns[1]][idx],saqd,idx,skuAndQtyDict[mainFileData.columns[2]][indx])
                    if (storeSkuAndQtyDict[store.columns[2]][idx]!=skuAndQtyDict[mainFileData.columns[3]][indx] and saqd==storeSkuAndQtyDict[store.columns[1]][idx]):
                        boo=True
                        storeSkuAndQtyDict[store.columns[2]][idx]=skuAndQtyDict[mainFileData.columns[3]][indx]
        df=pd.DataFrame(storeSkuAndQtyDict)         
        print(boo)
        
        
        if boo:
            writer = pd.ExcelWriter(path1, engine='xlsxwriter')
            df.to_excel(writer, sheet_name='Sheet1', index=False)
            writer.save()
            boo=False
    # all stores are updated
    messagebox.showinfo("FT","Stores are updated.")
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
    x = 347, y = 305,
    width = 307,
    height = 46)

# update all stores
img1 = PhotoImage(file = f"img1.png")
b1 = Button(
    image = img1,
    borderwidth = 0,
    highlightthickness = 0,
    command = updateAllStores,
    relief = "flat")

b1.place(
    x = 347, y = 431,
    width = 307,
    height = 46)

# choose stores location

img2 = PhotoImage(file = f"img2.png")
b2 = Button(
    image = img2,
    borderwidth = 0,
    highlightthickness = 0,
    command = ChooseStoreLocation,
    relief = "flat")

b2.place(
    x = 347, y = 369,
    width = 307,
    height = 46)


background_img = PhotoImage(file = f"background.png")
background = canvas.create_image(
    500.0, 216.0,
    image=background_img)

window.resizable(False, False)
window.mainloop()
