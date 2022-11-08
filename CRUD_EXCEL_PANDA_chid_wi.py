from tkinter import *
from tkinter import ttk,filedialog,messagebox
import pandas as pd 
import os

def main(): 
    win = Tk()
    app = Main_Win(win)
    win.mainloop()

class Main_Win:

    def __init__(self,root):     
        self.root = root
        self.root.title("Excel CRUD")
        self.root.geometry("800x600")    

        self.ID    = StringVar()
        self.first = StringVar()
        self.last  = StringVar()
        self.email = StringVar()
        
    
        Frame_1 = Frame(self.root, bg = 'grey',height= 500 , width = 600 )
        Frame_1.place(x= 10 , y= 10)    
        Button(Frame_1,text='new excel' ,command = self.new_excel   ).grid(row=5 ,column=0 ,pady=4)
        Button(Frame_1,text='Open excel',command = self.open_excel  ).grid(row=5 ,column=1 ,pady=4)
        Button(Frame_1,text='Exit'      ,command = self.root.quit   ).grid(row=5 ,column=6 ,pady=4)
    
        # ==================== treeView ==================================
        my_frame = Frame(self.root,height= 250 , width = 600)
        my_frame.place(x= 10 , y= 300)
    
        scroll_x = ttk.Scrollbar(my_frame,orient = HORIZONTAL)
        scroll_y = ttk.Scrollbar(my_frame,orient = VERTICAL  )
    
        self.my_tree = ttk.Treeview(my_frame,xscrollcommand=scroll_x.set,yscrollcommand=scroll_y.set)
    
        scroll_x.pack(side=BOTTOM,fill=X)
        scroll_y.pack(side=RIGHT,fill=Y)
    
        scroll_x.config(command = self.my_tree.xview)
        scroll_y.config(command = self.my_tree.yview)
        self.my_tree.pack(fill=BOTH,expand=1)

        self.my_tree.bind("<Button-3>", self.popup)
    
        self.my_tree.bind("<ButtonRelease-1>",self.get_cursor)   

        self.aMenu = Menu(self.root, tearoff=0)
        self.aMenu.add_command(label='New', command=lambda:self.go_New_Entry_Win())
        self.aMenu.add_command(label='Update', command=lambda:self.go_Update_Entry_Win())
        self.aMenu.add_command(label='delete', command=lambda: self.delete_Row())
        
    def open_excel(self):
    
        global wb_path , df 
        wb_path_Cur = filedialog.askopenfilename(initialdir= os.getcwd(),title= "Choose File",
                                                filetypes=(("Excel files", "*.xlsx")         ,
                                                           ("All files", "*.*")
                                                          )
                                                )
        wb_path = str(wb_path_Cur)
        df  = pd.read_excel(wb_path)
        
        self.ref_tree()
    
    def new_excel(self): 
        
        global wb_path , df 
        wb_path ='loc_Example.xlsx'
    
        if os.path.isfile(wb_path):
            pass
        else :
            path_Excel = filedialog.askdirectory(initialdir= os.getcwd(),title= "Output Excel")
            wb_path = str(path_Excel + '/' + 'loc_Example.xlsx')
            df = pd.DataFrame({'ID':[],'first':[] ,'last':[],'email':[]})
            df.to_excel(wb_path, index=False)  
    
        df = pd.read_excel(wb_path)
    
        self.ref_tree()

    def ref_tree(self):
      
        df  = pd.read_excel(wb_path)
        
        self.clear_tree()
        self.my_tree["column"] = list(df.columns)
        self.my_tree["show"]   = "headings" 
    
        for column in self.my_tree["column"]:
            self.my_tree.heading(column , text = column ,anchor=CENTER )
            self.my_tree.column (column , anchor=CENTER )   
        df_rows = df.to_numpy().tolist()
    
        for row in df_rows: 
            self.my_tree.insert("" , "end",values = row)
    
    def get_cursor(self,event=""):     
        cursor_row = self.my_tree.focus()
        
        content = self.my_tree.item(cursor_row)
        
        row = content["values"]
        
        self.first.set(row[1])
        self.last .set(row[2])
        self.email.set(row[3])
    
    def clear_tree(self): 

        self.my_tree.delete(*self.my_tree.get_children())

    def delete_Row(self):

        df = pd.read_excel(wb_path)
        selected_iid = self.my_tree.selection()[0]
        i = self.my_tree.index(selected_iid)
        
        df = df.drop(i)
        df = df.reset_index(drop=True)  
        
        for i in df.index :
           df.loc[i , ['ID'] ]= [i+1]

        df.to_excel(wb_path , index=False)

        self.clear_tree()
        self.ref_tree()
        
        messagebox.showinfo("Delete info","Emloyee Details Deleted Succesfully")

    def popup(self, event):
        """action in event of button 3 on tree view"""
        # select row under mouse
        iid = self.my_tree.identify_row(event.y)
        if iid:
            # mouse pointer over item
            self.my_tree.selection_set(iid)
            self.aMenu.post(event.x_root, event.y_root)            
        else:
            # mouse pointer not over item
            # occurs when items do not fill frame
            # no action required
            pass

    def go_New_Entry_Win(self): 
        self.new_window = Toplevel(self.root)
        self.app = New_Entry_Win(self.new_window)

    def go_Update_Entry_Win(self): 
        self.new_window = Toplevel(self.root)
        self.app = Update_Entry_Win(self.new_window)


class New_Entry_Win:

    def __init__(self,root):     
        self.root = root
        self.root.title("New Entries")
        self.root.geometry("500x600")    

        self.ID    = StringVar()
        self.first = StringVar()
        self.last  = StringVar()
        self.email = StringVar()

        Frame_1 = Frame(self.root, bg = 'grey',height= 500 , width = 600 )
        Frame_1.place(x= 10 , y= 10)

        self.first_Lbl = Label(Frame_1,text = "first")
        self.first_Lbl.grid(row=0,column=0)

        self.last_Lbl = Label(Frame_1,text = "last" )
        self.last_Lbl.grid(row=1,column=0)

        self.email_Lbl = Label(Frame_1,text = "email")
        self.email_Lbl.grid(row=2,column=0)

        self.first_En = Entry(Frame_1,textvariable=self.first)
        self.first_En.grid(row=0,column=1)

        self.last_En = Entry(Frame_1,textvariable=self.last)
        self.last_En.grid(row=1,column=1)

        self.email_En = Entry(Frame_1,textvariable=self.email)
        self.email_En.grid(row=2,column=1)

        Button(Frame_1,text='create',command = self.Create_Input).grid(row=3 ,column=1 ,pady=4)

    def return_Main(self): 

        self.root.destroy()
      

    def Create_Input(self):

        df  = pd.read_excel(wb_path)

        SeriesA = df['ID']
        SeriesB = df['first']
        SeriesC = df['last']
        SeriesD = df['email']

        self.ID = df.shape[0]+1
        
        A = pd.Series(self.ID)
        B = pd.Series(self.first_En.get())
        C = pd.Series(self.last_En .get())
        D = pd.Series(self.email_En.get())

        new_row = pd.DataFrame([A,B,C,D], index=['ID' ,'first','last','email']).T
        
        df_new = pd.concat((df, new_row), ignore_index= True) 
        df_new.to_excel(wb_path, index=False)

        Main_Win
        clr = Main_Win.clear_tree
        ref = Main_Win.ref_tree

        messagebox.showinfo("Create info","Data Created Succesfully")

        
        self.return_Main()

class Update_Entry_Win:

    def __init__(self,root):     
        self.root = root
        self.root.title("Update Entries")
        self.root.geometry("500x600")    

        self.ID    = StringVar()
        self.first = StringVar()
        self.last  = StringVar()
        self.email = StringVar()

        Frame_1 = Frame(self.root, bg = 'grey',height= 500 , width = 600 )
        Frame_1.place(x= 10 , y= 10)

        self.first_Lbl = Label(Frame_1,text = "first")
        self.first_Lbl.grid(row=0,column=0)

        self.last_Lbl = Label(Frame_1,text = "last" )
        self.last_Lbl.grid(row=1,column=0)

        self.email_Lbl = Label(Frame_1,text = "email")
        self.email_Lbl.grid(row=2,column=0)

        self.first_En = Entry(Frame_1,textvariable=self.first)
        self.first_En.grid(row=0,column=1)

        self.last_En = Entry(Frame_1,textvariable=self.last)
        self.last_En.grid(row=1,column=1)

        self.email_En = Entry(Frame_1,textvariable=self.email)
        self.email_En.grid(row=2,column=1)

        Button(Frame_1,text='Update',command = self.Update_Cells).grid(row=3 ,column=1 ,pady=4)

    def Update_Cells(self):
         
        df = pd.read_excel(wb_path)

        selected_iid = Main_Win.my_tree.selection()[0]
        i =   Main_Win.my_tree.index(selected_iid)

        new_First = self.first_En.get()
        new_Last  = self.last_En .get()
        new_email = self.email_En.get()

        df.loc[i , ['first','last','email']] = [new_First ,new_Last,new_email]
        
        df = df.reset_index(drop=True)
        df.to_excel(wb_path, index=False)

        clear_tr_1 = Main_Win.clear_tree
        ref_tr_1 = Main_Win.ref_tree
       
        messagebox.showinfo("Update info","Data Updated Succesfully")

        self.return_Main()        

    def return_Main(self): 

        self.root.destroy()



if __name__ == "__main__": 
   main() 