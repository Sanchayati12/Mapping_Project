#!/usr/bin/env python
# coding: utf-8

# 
# 1. refer : https://www.youtube.com/watch?v=XI_SSafLxCw
# 2. desktop app: refer : https://www.youtube.com/watch?v=QWqxRchawZY

# In[ ]:





# In[3]:


import tkinter as tk
from PIL import ImageTk, Image
# import PIL.Image
from tkinter import filedialog, messagebox, ttk
from tkinter import Toplevel, Button, Menu
from tkinter import *
from tkinter.filedialog import asksaveasfile
import xlwt
from xlwt import Workbook

from sklearn.feature_extraction.text import TfidfVectorizer
import numpy as np
from sklearn.feature_extraction import text


import pandas as pd
import pickle

# initalise the tkinter GUI

root = tk.Tk()
root.iconbitmap("C:\\Users\\Sanchayati.Bajganiya\\.vscode\\ML Project\\logo1.ico") #
blank_space =" " # One empty space
#root = tk.Tk(className = "Mapping using ML Model")
root.title(150*blank_space+'MAPPING USING ML MODEL')


#backgroundImage

# img = ImageTk.PhotoImage(Image.open("lm3JuT.jpg"))  
img = ImageTk.PhotoImage(file="lm3JuT.jpg") 
l=Label(image=img)
l.pack()

  
#root.geometry("500x500")
root.geometry("1199x600+100+50") # set the root dimensions
root.pack_propagate(False) # tells the root to not let the widgets inside it determine its size.
root.resizable(0, 0) # makes the root window fixed in size.

# Frame for TreeView
frame1 = tk.LabelFrame(root, text="Ultrasound Mapping Q1 2021",font='Helvetica 18 bold',bg='orange')
#frame1.place(height=250, width=500)
frame1.place(x=0,y=1,height=400, width=1200)

# Frame for open file dialog
file_frame = tk.LabelFrame(root, text="Open File",fg='orange',font='Helvetica 10 bold')
#file_frame.place(height=100, width=400, rely=0.65, relx=0)
file_frame.place(x=0,y=360,height=70, width=1200) #

# Buttons
button1 = tk.Button(file_frame, text="Browse A File", command=lambda: File_dialog(),fg='Black',font='Helvetica 8 bold')
button1.place(x=500, y=20)

button2 = tk.Button(file_frame, text="Load File", command=lambda: Load_excel_data(),fg='Black',font='Helvetica 8 bold')
button2.place(x=1000, y=20) #


button3= tk.Button(file_frame,text="Save File", command=lambda: save_data(),fg='Black',font='Helvetica 8 bold')
# button3.place(rely=0.65,relx=0.10)
button3.place(x=100,y=20) 

# The file/file path text
label_file = ttk.Label(file_frame, text="No File Selected")
label_file.place(rely=0, relx=0)

## Treeview Widget
tv1 = ttk.Treeview(frame1)
tv1.place(relheight=1, relwidth=1) # set the height and width of the widget to 100% of its container (frame1).

treescrolly = tk.Scrollbar(frame1, orient="vertical", command=tv1.yview) # command means update the yaxis view of the widget
treescrollx = tk.Scrollbar(frame1, orient="horizontal", command=tv1.xview) # command means update the xaxis view of the widget
tv1.configure(xscrollcommand=treescrollx.set, yscrollcommand=treescrolly.set) # assign the scrollbars to the Treeview Widget
treescrollx.pack(side="bottom", fill="x") # make the scrollbar fill the x axis of the Treeview widget
treescrolly.pack(side="right", fill="y") # make the scrollbar fill the y axis of the Treeview widget


def File_dialog():
    """This Function will open the file explorer and assign the chosen file path to label_file"""
    filename = filedialog.askopenfilename(initialdir="/",
                                          title="Select A File",
                                          filetype=(("xlsx files", "*.xlsx"),("All Files", "*.*")))
    label_file["text"] = filename
    return None


def Load_excel_data():
    """If the file selected is valid this will load the file into the Treeview"""
    #Pkl_Filename = "USG_saved_model.pkl"
    Pkl_Filename = "Final_USG_model.pkl"
    global df
    
    # Load the Model back from file
    with open(Pkl_Filename, 'rb') as file:  
        Pickled_USG_Model = pickle.load(file)
        
    file_path = label_file["text"]
    try:
        excel_filename = r"{}".format(file_path)
        if excel_filename[-4:] == ".csv":
            df = pd.read_csv(excel_filename, sheetname='data')
            similarity_data = pd.read_csv(excel_filename, sheetname='similarity_data')
            
            N_results = df["Product Description"]
            N_predictions = Pickled_USG_Model.predict(N_results)
            df["Machine/Parts/Others"] = N_predictions
            
            
            main_corpus = []
            for i in similarity_data["Product Description"]: #
                i = str(i)
                main_corpus.append(i)
                
            df_mt = df[df["Machine/Parts/Others"] == "Machine"]
            
            my_stop_words = text.ENGLISH_STOP_WORDS.union(["Ultrasonic","Scanning","Apparatus","ULTRASOUND","SYSTEM","APPARATUS","ULTRASONIC",
                                               "ACC","ACCESS","WITH","MEDICAL","EQUIPMENT","SYS","STANDARD","ACCESSORIES",
                                              "Electro-Diagnostic","COLOR","DOPPLER","DIAGNOSTIC","Other","BRAND","NEW","MEDICAL","EQUIPMENT",
                                              "SCANNER","STD","W.STD.ACC.","PORTABLE","UNIT","MAIN",])
            corpus1 = main_corpus
            vect = TfidfVectorizer(stop_words=my_stop_words, ngram_range = (1,2))
            req_index = []

            for i in df_mt["Product Description"]:
                corpus1.append(i)
                tfidf = vect.fit_transform(corpus1)
                pairwise_similarity = tfidf * tfidf.T

                arr = pairwise_similarity.toarray()
                np.fill_diagonal(arr, np.nan)

                input_doc = str(i)
                print("Input sentence == ",i)
                input_idx = corpus1.index(input_doc)
                #print("Index of input sentence, should be in last ==",input_idx)
                print()

                result_idx = np.nanargmax(arr[3775])
                req_index.append(result_idx)
                print("Input sentence mapped to output == ",corpus1[result_idx])
                print("req_index ==",len(req_index))
                print("XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX")
                print()

                del corpus1[-1]
            
            
            
            result = similarity_data.iloc[req_index]
            result_1 = result.reset_index(drop=True)
            
            change_index = []
            for i in data_mt.index:
                change_index.append(i)
                
            s = pd.Series(change_index)
            match = result_1.set_index([s])
            
            df["Type"] = match["Type"]
            df["Used/New"] = match["Used/New"]
            df["Company"] = match["Company"]
            df["Brand"] = match["Brand"]

    
                
            
            
            
        else:
            df = pd.read_excel(excel_filename,sheet_name='data')
            similarity_data = pd.read_excel(excel_filename, sheet_name='similarity_data')
            
            N_results = df["Product Description"]
            N_predictions = Pickled_USG_Model.predict(N_results)
            df["Machine/Parts/Others"] = N_predictions
            
            
            main_corpus = []
            for i in similarity_data["Product Description"]:
                i = str(i)
                main_corpus.append(i)
                
            df_mt = df[df["Machine/Parts/Others"] == "Machine"]
            
            my_stop_words = text.ENGLISH_STOP_WORDS.union(["Ultrasonic","Scanning","Apparatus","ULTRASOUND","SYSTEM","APPARATUS","ULTRASONIC",
                                               "ACC","ACCESS","WITH","MEDICAL","EQUIPMENT","SYS","STANDARD","ACCESSORIES",
                                              "Electro-Diagnostic","COLOR","DOPPLER","DIAGNOSTIC","Other","BRAND","NEW","MEDICAL","EQUIPMENT",
                                              "SCANNER","STD","W.STD.ACC.","PORTABLE","UNIT","MAIN",])
            
            corpus1 = main_corpus
            vect = TfidfVectorizer(stop_words=my_stop_words, ngram_range = (1,2))
            req_index = []

            for i in df_mt["Product Description"]:
                corpus1.append(i)
                tfidf = vect.fit_transform(corpus1)
                pairwise_similarity = tfidf * tfidf.T

                arr = pairwise_similarity.toarray()
                np.fill_diagonal(arr, np.nan)

                input_doc = str(i)
                print("Input sentence == ",i)
                input_idx = corpus1.index(input_doc)
                #print("Index of input sentence, should be in last ==",input_idx)
                print()

                result_idx = np.nanargmax(arr[3775])
                req_index.append(result_idx)
                print("Input sentence mapped to output == ",corpus1[result_idx])
                print("req_index ==",len(req_index))
                print("XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX")
                print()

                del corpus1[-1]
                
            result = similarity_data.iloc[req_index]
            result_1 = result.reset_index(drop=True)
            
            change_index = []
            for i in df_mt.index:
                change_index.append(i)
                
            s = pd.Series(change_index)
            match = result_1.set_index([s])
            
            df["Type"] = match["Type"]
            df["Used/New"] = match["Used/New"]
            df["Company"] = match["Company"]
            df["Brand"] = match["Brand"]
            

            
            

#             savefile = filedialog.asksaveasfilename(filetypes=(("Excel files", "*.xlsx"),
#                                                          ("All files", "*.*") ))               
#              data = pd.read_excel(file,sheetname="Sheet1")
#             df.to_excel(savefile + ".xlsx", index=False, sheet_name="Results")

    except ValueError:
        tk.messagebox.showerror("Information", "The file you have chosen is invalid")
        return None
    except FileNotFoundError:
        tk.messagebox.showerror("Information", f"No such file as {file_path}")
        return None

    clear_data()
    tv1["column"] = list(df.columns)
    tv1["show"] = "headings"
    for column in tv1["columns"]:
        tv1.heading(column, text=column) # let the column heading = column name

    df_rows = df.to_numpy().tolist() # turns the dataframe into a list of lists
    for row in df_rows:
        tv1.insert("", "end", values=row) # inserts each list into the treeview. For parameters see https://docs.python.org/3/library/tkinter.ttkhtml#tkinter.ttk.Treeview.insert
    return None

def clear_data():
    tv1.delete(*tv1.get_children())
    return None
def save_data():
    savefile = filedialog.asksaveasfilename(filetypes=(("Excel files", "*.xlsx"),
                                                         ("All files", "*.*") ))               
    df.to_excel(savefile + ".xlsx", index=False, sheet_name="Results")


root.mainloop()


# In[ ]:




