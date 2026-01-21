from tkinter import *
from tkinter import messagebox
from tkinter import ttk
import Backend
import random
import getpass
from openpyxl import load_workbook, Workbook
POS = Tk()
POS.geometry("1100x600+150+10")
loginframe = Frame(POS, background='#275c70')
salesframe = Frame(POS,background="#4D6E55")

Sales_ID = random.randint(1000,9999)
Temp_Total = 0.0
total_var = StringVar()
total_var.set(f"Total:{Temp_Total:.2f} ")
change = 0.0
is_admin = False
is_audit = False

#----------------------LOGIN FRAME-------------------------------------------------------------------
for i in range(11):
    loginframe.rowconfigure(i,weight=1)
for i in range(10):
    loginframe.columnconfigure(i,weight=2)

label = Label(loginframe,
              text="POSSYS SOLUTIONS",
              font=("times new roman",24,'bold'),
              bg="#00403d",
              fg='#B8E3E9',
              relief=RIDGE,
              bd = '15',
              width= 20)
label.place(relx=0.52, rely=0.1, width = 600, anchor='center')

label = Label(loginframe,
              text="Final Capstone Project: Sales and Inventory Management System",
              font=("times new roman",14,'bold'),
              bg="#00403d",
              fg='#B8E3E9',
              relief=RIDGE,
              bd = '15',
              width= 20)
label.place(relx=0.52, rely=0.2, width=600, anchor='center')

user_entry = Entry(loginframe,font=('Arial', 20, 'bold'), width=30,)
Label(loginframe,text="USERNAME",
              font=("times new roman",17),
              bg="#00403d",
              fg='#B8E3E9',
              relief=RIDGE,
              bd = '10').grid(row=4,column=4)
user_entry.grid(row=4,column=5,sticky="ew")

Label(loginframe, text="Password",
              font=("times new roman",19),
              bg="#00403d",
              fg='#B8E3E9',
              relief=RIDGE,
              bd = '10',
              width=9).grid(row=5,column=4)
user_password = Entry(loginframe,font=('Arial', 20, 'bold'), width=30, show='*')
user_password.grid(row=5,column=5,sticky="ew")


def login_user():
    global is_admin
    global is_audit
    if Backend.login(user_entry.get(),user_password.get()):
        if Backend.checkadmin(user_entry.get()):
            is_admin = True
        elif Backend.checkaudit(user_entry.get()):
            is_audit = True
        user_entry.delete(0,END)
        user_password.delete(0,END)
        loginframe.pack_forget()
        salesframe.pack(fill = "both",expand=True)
    else:
        warn = Label(loginframe,text = "!INVALID Username or Password!",
              font=("times new roman",17,'bold'),
              bg="#00403d",
              fg='#B8E3E9',
              relief=RAISED,
              bd = '5',
              width=26)
        warn.grid(row=7,column=5)

Button(loginframe, text = 'Login',font=("times new roman",17,'bold'),
              bg="#00403d",
              fg='#B8E3E9',
              activebackground='#4bafd6',
              activeforeground='white',
              relief=GROOVE,
              bd = '10',command = login_user).grid(row=7,column=4,sticky="news")
loginframe.pack(fill = "both",expand=True)
devs = Label(loginframe,
             text='Developed by: \n-JUSTINE BONGALOS,\n-JOAQUIN DORONGON,\n-MARK OCAMPO,\n -MJ DAMIAN',
             font=('Arial', 10, 'bold'),
              bg="#00403d",
              fg='#B8E3E9',
             relief=SOLID,
             bd=5)
devs.grid(row=10,column=9)

#------------------------MAIN POS FRAME-----------------------------------------------------
for i in range(10):
    salesframe.rowconfigure(i,weight=1)
for i in range(7):
    salesframe.columnconfigure(i,weight=1)
Label(salesframe,text="POSSYS SOLUTIONS",
      font=('times new roman',24,'bold'),
      bg="#00403d",
      fg='#B8E3E9',
      relief=RIDGE,
      bd=10).grid(row=0,column=4,columnspan=1,
      sticky='nswe')

product_entry = Entry(salesframe)
Label(salesframe,text="ENTRY >>>",font=("Times new Roman",15,"bold")).grid(row=1,column=1,sticky="we")
product_entry.grid(row=1,column=3,columnspan=3,sticky="we")
item_list = ttk.Treeview(salesframe,columns=("Product","Quantity","Price","Subtotal"),show="headings",height=10)
item_list.grid(row=2,column=3,columnspan=3,rowspan=7,sticky='nsew')
item_list.heading("Product",text="Product")
item_list.heading("Quantity",text="Quantity")
item_list.heading("Price",text="Price")
item_list.heading("Subtotal",text="Subtotal")
def addtocart(event=None):
    global Temp_Total
    product = product_entry.get()
    if "*" in product_entry.get():
        try:
            product_code, qty = product.split("*")
            product_code = product_code.strip()
            qty = int(qty.strip())
        except:
            messagebox.showerror("Invalid entry format!", "Use: Code * Quantity")
            product_entry.delete(0, END)
            return
    else:
        product_code = product
        qty = 1
    
    if Backend.get_product(product_code):
        Backend.buy(product_code,qty,Sales_ID)
        item_list.insert("",END,values=(Backend.product_name(product_code),qty,Backend.get_price(product_code),float(qty*Backend.get_price(product_code))))
        product_entry.delete(0, END)
        Temp_Total += float(qty*Backend.get_price(product_code))
        total_var.set(f"Total: {Temp_Total:.2f}")
    else:
        messagebox.showerror("WARNING","INVALID PRODUCT ID")
        product_entry.delete(0, END)
def finish_transaction():
    popup = Toplevel(POS)
    popup.title("Finish Transaction")
    popup.geometry("250x180")
    popup.grab_set()
    Label(popup,text = f"TOTAL: {Temp_Total}").pack()
    Label(popup,text="Cash Recieved:").pack()
    cash = Entry(popup)
    cash.pack()
    def changes(event=None):
        global Temp_Total
        global change
        global Sales_ID
        try:
            change += (float(cash.get()) - Temp_Total)
            messagebox.showinfo("CHANGE",f" Your Change is: {change}")
            Sales_ID = random.randint(1000,9999)
            for item in item_list.get_children():
                item_list.delete(item)
            Temp_Total = 0
            total_var.set(f"Total: {Temp_Total:.2f} ")
            Backend.save()
            popup.destroy()


        except ValueError:
            messagebox.showerror("Invalid entry format!", "YOU CANNOT SUBTRACT THAT TO A NUMBER YEA?")
            cash.delete(0, END)

    Button(popup,text="Confirm",command=changes).pack()

def printreceipt():
    popup = Toplevel(POS)
    popup.title("PRINT RECEIPT")
    popup.geometry("700x500")
    popup.grab_set()
    Label(popup,text="Enter Trasaction ID").pack()
    transac_id = Entry(popup)
    transac_id.pack()
    item_list = ttk.Treeview(popup,columns=("Product","Quantity","Price","Subtotal"),show="headings",height=10)
    item_list.pack()
    def printrec():
        wb2 = load_workbook("sale.xlsx")
        ws2 = wb2.active
        for row in ws2.iter_rows(min_row=1,values_only=True):
            if row[0] == int(transac_id.get().strip()):
                item_list.insert("",END,values=(row[2], row[3], row[4], row[5]))
            
    def close():
        popup.destroy()
    Button(popup,text="Submit",
      font=('times new roman',10,'bold'),
      bg="#00403d",
      fg='#B8E3E9',
      relief=RIDGE,
      bd=10,
      command=printrec).pack()
    Button(popup,text="Close",
      font=('times new roman',10,'bold'),
      bg="#00403d",
      fg='#B8E3E9',
      relief=RIDGE,
      bd=10,
      command=close).pack()

def logout():
    pop = Toplevel(POS)
    pop.title("WARNING")
    pop.geometry("300x200")
    Label(pop,text="LOG OUT?",font=("times new roman",18,"bold")).pack()
    def yeah():
        pop.destroy()
        salesframe.pack_forget()
        loginframe.pack(fill = "both",expand=True)
    def nah():
        pop.destroy()
    Button(pop,text="Yes",command=yeah).pack()
    Button(pop,text="Nah",command=nah).pack()

def change_stock():
    if is_audit or is_admin:
        Inventory_Database = "Database.xlsx"
        wb1 = load_workbook(Inventory_Database)
        ws1 = wb1.active
        popup = Toplevel(POS)
        popup.title("Change stock")
        popup.geometry("700x500")
        popup.grab_set()
        Label(popup,text="Low Stocks",font=("Times New Roman",15,"bold")).pack()
        stocks_table = ttk.Treeview(popup,columns=("id","name","stock","order"),show="headings")
        stocks_table.heading("id", text="Rank")
        stocks_table.heading("name", text="Product Name")
        stocks_table.heading("stock", text="Stock/s")
        stocks_table.heading("order", text="ORDER")
        stocks_table.column("id", width=60, anchor=CENTER)
        stocks_table.column("name",width=100)
        stocks_table.column("stock", width=200, anchor=W)
        stocks_table.column("order", width=100, anchor=CENTER)
        stocks = Backend.low_stock_alerts()
        for products in stocks:
            if products["product_id"] != None:
                stocks_table.insert(
                    "",
                    END,
                    values=(
                    products["product_id"],
                    products["name"],
                    products["stock"],
                    products["reorder"]
                    )
                )
        stocks_table.pack()
        Label(popup,text="Enter Product ID").pack()
        product_entry = Entry(popup)
        product_entry.pack()
        Label(popup,text="Enter New Stock").pack()
        stock_entry = Entry(popup)
        stock_entry.pack()
        def confirmstock():
            if Backend.get_product(product_entry.get().strip().upper()):
                Backend.change_stock(product_entry.get().strip(),stock_entry.get().strip())
                messagebox.showinfo("STOCKS","Stocks has been updated")
                popup.destroy()
            else:
                messagebox.showwarning("WARNING","ID NOT FOUND")
        def close():
            popup.destroy()
        Button(popup,text="Submit",command=confirmstock).pack()
        Button(popup,text="Cancel",command=close).pack()
    else:
        messagebox.showwarning("WARNING","You Need to be an AUDIT/ADMIN to make changes")
def add_product():
    if is_audit or is_admin:
        popup = Toplevel(POS)
        popup.title("Add Product")
        popup.geometry("700x500")
        popup.grab_set()
        Label(popup,text="Enter Product ID").pack()
        product = Entry(popup)
        product.pack()
        Label(popup,text="Enter Product Name").pack()
        name = Entry(popup)
        name.pack()
        Label(popup,text="Enter Product Category").pack()
        category = Entry(popup)
        category.pack()
        Label(popup,text="Enter Product Price").pack()
        price = Entry(popup)
        price.pack()
        Label(popup,text="Enter Product Stock/s").pack()
        stocks = Entry(popup)
        stocks.pack()
        Label(popup,text="Enter Product Reorder lvl").pack()
        reorder = Entry(popup)
        reorder.pack()
        def submit():
            try:
                Backend.add_new(product.get(),name.get(),category.get(),price.get(),stocks.get(),reorder.get())
                messagebox.showinfo("SUCESS",f"Successfully added {name.get()} worth {price.get()} with {stocks.get()} stocks")
                product.delete(0,END)
                name.delete(0,END)
                category.delete(0,END)
                price.delete(0,END)
                stocks.delete(0,END)
                reorder.delete(0,END)
                
            except ValueError:
                messagebox.showwarning("Warning", "pls enter integers/decimals on price,stocks and reorder lvl")
                price.delete(0,END)
                stocks.delete(0,END)
                reorder.delete(0,END)
        def exit():
            popup.destroy()
        Button(popup,text="Submit",command=submit).pack()
        Button(popup,text="Exit",command=exit).pack()
    else:
        messagebox.showwarning("WARNING","You Need to be an AUDIT/ADMIN to make changes")
def remove():
    if is_admin or is_audit:
        popup = Toplevel(POS)
        popup.title("Remove Product")
        popup.geometry("200x200")
        popup.grab_set()
        Label(popup,text="Enter Product ID to be Removed").pack()
        product = Entry(popup)
        product.pack()
        def submit():
            if Backend.get_product(product.get()):
                pop = Toplevel(POS)
                pop.title("Confirm Removal")
                pop.geometry("250x150")
                pop.grab_set()
                Label(pop,text=f"Are you sure you wanna remove {Backend.product_name(product.get())}?").pack()
                def ye():
                    pop.destroy()
                    messagebox.showinfo("SUCCESS",f"Sucessfully removed {Backend.product_name(product.get())} ")
                    Backend.remove(product.get())
                    product.delete(0,END)
                def nah():
                    pop.destroy()

                Button(pop,text="YE",command=ye).pack()
                Button(pop,text="NAH",command=nah).pack()
            else:
                messagebox.showwarning("W","Cant Find Code")
        def cancel():
            popup.destroy()
        Button(popup,text="Submit",command=submit).pack()
        Button(popup,text="Cancel",command=cancel).pack()
    else:
        messagebox.showwarning("WARNING","You Need to be an AUDIT/ADMIN to make changes")    

def sales_summary():
    popup = Toplevel(POS)
    popup.title("ANALYTICS")
    popup.geometry("550x250")
    popup.grab_set()
    Label(popup,text="PRODUCT RANKINGS BASED ON NO# of SALES",font=("times new roman",16,"bold")).pack()
    rank_table = ttk.Treeview(popup,columns=("rank","id","product","sold"),show="headings")
    rank_table.heading("rank", text="Rank")
    rank_table.heading("product", text="Product Name")
    rank_table.heading("sold", text="Units Sold")
    rank_table.heading("id",text="Product ID")
    rank_table.column("rank", width=60, anchor=CENTER)
    rank_table.column("id",width=100)
    rank_table.column("product", width=200, anchor=W)
    rank_table.column("sold", width=100, anchor=CENTER)
    rankings = Backend.best_selling_products()
    rank_table.pack()
    for rank, item in enumerate(rankings, start=1):
        rank_table.insert(
            "",
            "end",
            values=(
            rank,
            item["product_id"],
            item["name"],
            item["total_sold"])
        )

def remove_user():
    if is_admin:
        popup = Toplevel(POS)
        popup.title("Remove User")
        popup.geometry("200x200")
        popup.grab_set()
        Label(popup,text="Enter Username to be Removed").pack()
        user = Entry(popup)
        user.pack()
        def submit():
            pop = Toplevel(POS)
            pop.title("Confirm Removal")
            pop.geometry("250x150")
            pop.grab_set()
            Label(pop,text=f"Are you sure you wanna remove {user.get()}?").pack()
            def ye():
                pop.destroy()
                messagebox.showinfo("SUCCESS",f"Sucessfully removed {user.get()} ")
                Backend.remove_ppl(user.get())
                user.delete(0,END)
            def nah():
                pop.destroy()

            Button(pop,text="YE",command=ye).pack()
            Button(pop,text="NAH",command=nah).pack()
        def cancel():
            popup.destroy()
        Button(popup,text="Submit",command=submit).pack()
        Button(popup,text="Cancel",command=cancel).pack()
    else:
        messagebox.showwarning("WARNING","You Need to be an ADMIN to make changes")
def add_user():
    if is_admin:
        roles = ["Admin","Audit","Cashier"]
        selected_role = StringVar()
        selected_role.set(roles[0])
        popup = Toplevel(POS)
        popup.title("Add People")
        popup.geometry("700x500")
        popup.grab_set()
        Label(popup,text="Enter Username").pack()
        user = Entry(popup)
        user.pack()
        Label(popup,text="Enter Password").pack()
        password = Entry(popup)
        password.pack()
        role_menu = OptionMenu(popup, selected_role, *roles)
        role_menu.pack(pady=20)
        def confirm():
            Backend.add_ppl(user.get(),role_menu.get(),password.get())
            messagebox.showinfo("SUCCESS",f"{user.get()} is sucessfully added as {role_menu.get()}")
        def cancel():
            popup.destroy()
        Button(popup,text="Add",command=confirm).pack()
        Button(popup,text="Cancel",command=cancel)
    else:
        messagebox.showwarning("WARNING","You Need to be an ADMIN to make changes")

product_entry.bind("<Return>",addtocart)
Label(salesframe,textvariable=total_var,font=("arial",18),
      bg="#00403d",
      fg='#B8E3E9',
      relief=GROOVE,
      bd = '5').grid(column=4,row=9,sticky="nsew",pady=10)
logout_button = Button(salesframe,text="Log Out",font=("arial",13),command=logout,
              bg="#00403d",
              fg='#B8E3E9',
              activebackground='#4bafd6',
              activeforeground='white',
              relief=GROOVE,
              bd = '5')
logout_button.grid(row=2,column=0,sticky="nsew",columnspan=2)


add_product_button = Button(salesframe,text="Add Product",font=("arial",13),command=add_product,
              bg="#00403d",
              fg='#B8E3E9',
              activebackground='#4bafd6',
              activeforeground='white',
              relief=GROOVE,
              bd = '5')
add_product_button.grid(row=3,column=0,sticky="nsew")


remove_product_button = Button(salesframe,text="Remove Product",font=("arial",13),command=remove,
              bg="#00403d",
              fg='#B8E3E9',
              activebackground='#4bafd6',
              activeforeground='white',
              relief=GROOVE,
              bd = '5')
remove_product_button.grid(row=4,column=0,sticky="nsew")


change_stock_button = Button(salesframe,text="Change Stock",font=("arial",13),command=change_stock,
              bg="#00403d",
              fg='#B8E3E9',
              activebackground='#4bafd6',
              activeforeground='white',
              relief=GROOVE,
              bd = '5')
change_stock_button.grid(row=5,column=0,sticky='nsew')


print_reciept_button = Button(salesframe,text="Print reciept",font=("arial",13),command=printreceipt,
              bg="#00403d",
              fg='#B8E3E9',
              activebackground='#4bafd6',
              activeforeground='white',
              relief=GROOVE,
              bd = '5')
print_reciept_button.grid(row=6,column=0,sticky='nsew')


sale_summary_button = Button(salesframe,text="Sales Summary",command=sales_summary,font=("arial",13),
              bg="#00403d",
              fg='#B8E3E9',
              activebackground='#4bafd6',
              activeforeground='white',
              relief=GROOVE,
              bd = '5')
sale_summary_button.grid(row=3,column=1,sticky='nsew')


add_user_button = Button(salesframe,text="Add User",command=add_user,font=("arial",13),
              bg="#00403d",
              fg='#B8E3E9',
              activebackground='#4bafd6',
              activeforeground='white',
              relief=GROOVE,
              bd = '5')
add_user_button.grid(row=4,column=1,sticky='nsew')


remove_user_button = Button(salesframe,text="Remove User",command=remove_user,font=("arial",13),
              bg="#00403d",
              fg='#B8E3E9',
              activebackground='#4bafd6',
              activeforeground='white',
              relief=GROOVE,
              bd = '5')
remove_user_button.grid(row=5,column=1,sticky='nsew')


finish_transac_button = Button(salesframe,text="Finish Transaction",font=("arial",13),command=finish_transaction,
              bg="#00403d",
              fg='#B8E3E9',
              activebackground='#4bafd6',
              activeforeground='white',
              relief=GROOVE,
              bd = '5')
finish_transac_button.grid(row=6,column=1,sticky="nsew")


POS.mainloop()