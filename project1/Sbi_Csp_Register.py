from tkinter import *
from tkinter import ttk
import mysql.connector as mys
import datetime
import time
from datetime import date
from tkinter import messagebox
from openpyxl import Workbook
from openpyxl.styles import Alignment
from openpyxl.drawing.image import Image
from openpyxl.styles import Font
from openpyxl.styles import Border,Side
import win32api
import win32print
import tkinter.messagebox
#------------------------------------------MAIN_WINDOW---------------------------------------------------------
root = Tk()
root.title("SBI CSP REGISTER")
root.geometry('1600x500')
root.configure(bg="dark blue")
photto = PhotoImage(file = "app_logo1.png")
root.iconphoto(False, photto)
#------------------------------------------EXTRAS---------------------------------------------------------
frame1=LabelFrame(root,relief=RIDGE,bd=10,bg="cyan")
frame1.place(x="70",y="50")

mainl=Label(frame1,text="SBI Customer Service Point",font=("Arial",36))
mainl.pack()

sb_2 = Scrollbar(root)
sb_2.place(x="1150",y="150")


#-------------------------------------------LABEL--------------------------------------------------------
l1=Label(root,text="MOB. NO:",font=("Arial",16),relief="raised",width="10",cursor="spider")
l1.place(x="75",y="150")

l2=Label(root,text="A/C NO:",font=("Arial",16),relief="raised",width="10",cursor="spider")
l2.place(x="75",y="200")

l21=Label(root,text="Amount",font=("Arial",16),relief="raised",width="10",cursor="spider")
l21.place(x="75",y="250")

l3=Label(root,text="BANK NAME:",font=("Arial",16),relief="raised",width="10",cursor="spider")
l3.place(x="75",y="300")

l4=Label(root,text="NAME:",font=("Arial",16),relief="raised",width="10",cursor="spider")
l4.place(x="75",y="350")


#--------------------------------------------ENTERY_TYPE-------------------------------------------------------


name_var=StringVar()
ac_int=StringVar()
bankai=StringVar()
mob_int=StringVar()
amt_int=StringVar()

#---------------------------------------------ENTRY_AND_COMBOBOX------------------------------------------------------
e1=Entry(root,textvariable=mob_int,font=("Arial",16),width="30")
e1.place(x="250",y="150")

trtype=['AEPS(withdraw)','AEPS(deposit)','ATM(deposit)','ATM(withdraw)']

lbox =ttk.Combobox(root,font = ('arial',16,'bold'), width = 30, textvariable =ac_int)
lbox['values'] = trtype
lbox.place(x="250",y="200")

e3=Entry(root,textvariable=amt_int,font=("Arial",16),width="30")
e3.place(x="250",y="250")

items=['Bank of Baroda','Bank of India','Bank of Maharashtra'
                                   ,'Canara Bank','Central Bank of India','Indian Bank','Indian Overseas Bank','Punjab and Sind Bank'
                                   ,'Punjab & Sind Bank','Punjab National Bank','State Bank of India','UCO Bank','Union Bank of India','Axis Bank Ltd.','Bandhan Bank Ltd.'
                                   ,'CSB Bank Ltd.','City Union Bank Ltd.','DCB Bank Ltd.','Dhanlaxmi Bank Ltd.','Federal Bank Ltd.','HDFC Bank Ltd']

lbox1 =ttk.Combobox(root,font = ('arial',16,'bold'), width = 30, textvariable =bankai)
lbox1['values'] = items
lbox1.place(x="250",y="300")

e5=Entry(root,textvariable=name_var,font=("Arial",16),width="30")
e5.place(x="250",y="350")

#------------------------------------------NEW_WINDOW---------------------------------------------------------
def open1():
    newwindow=Toplevel(root)
    newwindow.geometry('400x400')
    newwindow.configure(bg="dark blue")
    
    l11=Label(newwindow,text="500:",font=("Arial",12),relief="raised",cursor="spider")
    l11.place(x="50",y="50")

    l12=Label(newwindow,text="200:",font=("Arial",12),relief="raised",cursor="spider")
    l12.place(x="50",y="100")

    l13=Label(newwindow,text="100:",font=("Arial",12),relief="raised",cursor="spider")
    l13.place(x="50",y="150")

    l14=Label(newwindow,text="50:  ",font=("Arial",12),relief="raised",cursor="spider")
    l14.place(x="50",y="200")

    l15=Label(newwindow,text="20:  ",font=("Arial",12),relief="raised",cursor="spider")
    l15.place(x="50",y="250")

    l16=Label(newwindow,text="10:  ",font=("Arial",12),relief="raised",cursor="spider")
    l16.place(x="50",y="300")

    e500=StringVar()
    e200=StringVar()
    e100=StringVar()
    e50=StringVar()
    e20=StringVar()
    e10=StringVar()
    tot=StringVar()
    

    e11=Entry(newwindow,textvariable=e500,font=("Arial",12))
    e11.place(x="100",y="50")

    e12=Entry(newwindow,textvariable=e200,font=("Arial",12))
    e12.place(x="100",y="100")

    e13=Entry(newwindow,textvariable=e100,font=("Arial",12))
    e13.place(x="100",y="150")

    e14=Entry(newwindow,textvariable=e50,font=("Arial",12))
    e14.place(x="100",y="200")

    e15=Entry(newwindow,textvariable=e20,font=("Arial",12))
    e15.place(x="100",y="250")

    e16=Entry(newwindow,textvariable=e10,font=("Arial",12))
    e16.place(x="100",y="300")

    e17=Entry(newwindow,textvariable=tot,font=("Arial",12))
    e17.place(x="100",y="350")



    def summ():
        q500=int(e500.get() or 0)
        q200=int(e200.get() or 0)
        q100=int(e100.get() or 0)
        q50=int(e50.get() or 0)
        q20=int(e20.get() or 0)
        q10=int(e10.get() or 0)

        

        res1=q500*500
        res2=q200*200
        res3=q100*100
        res4=q50*50
        res5=q20*20
        res6=q10*10
        res=res1+res2+res3+res4+res5+res6

        tot.set(res)

    def reset2():

            e500.set("")
            e200.set("")
            e100.set("")
            e50.set("")
            e20.set("")
            e10.set("")
            tot.set("")

    def check1():
        amount=amt_int.get()
        if amount==res:
            tkinter.messagebox.showinfo(title="AMOUNT MISMATCH", message=None, **options)
        else:
            print("lala")
    
    b55=Button(newwindow,text="=",font=("Arial",12,'bold'),activebackground="cyan",relief="ridge",command=summ)
    b55.place(x="50",y="350")

    b66=Button(newwindow,text="Reset",font=("Arial",12,'bold'),activebackground="cyan",relief="ridge",command=reset2)
    b66.place(x="320",y="350")

    b77=Button(newwindow,text="Check",font=("Arial",12,'bold'),activebackground="cyan",relief="ridge",command=check1)
    b77.place(x="320",y="300")

#--------------------------------------------USER_DEF_FUNCTION-------------------------------------------------------
def print1():
    mob_no_r=mob_int.get()
    ac_no_r=ac_int.get()
    ac_name_r=name_var.get()
    bank_name_r=bankai.get()
    amount_r=amt_int.get()
    today1 = date.today()
    date2=today1.strftime("%d-%m-%Y")
    current_time = time.strftime('%H:%M:%S:%p')
    data = [
        ['','                           ग्राहक सेवा बिंदु |CUSTOMER SERVICE POINT',''],
        ['','                         வாடிக்கையாளர் சேவை மையம்',''],
        [ '                NO.40/123 THIRUNEERMALAI ROAD, NAGALKENI (NEAR MGR STATUE)', ''],
        [ 'CHROMEPET, CHENNAI-600 044', 'TEL : +91 44486 50739'],
        ['DATE :',date2+'          |          '+current_time, ''],
        ['MOBILE NO :',mob_no_r, ''],
        ['ACCOUNT NO :',ac_no_r, ''],
        ['ACCOUNT NAME :', ac_name_r],
        ['BANK NAME :',bank_name_r],
        ['TRF. AMOUNT :', amount_r]
    ]

    filename = 'customer_data.xlsx'

    
    
    # Create a new workbook
    workbook = Workbook()
    sheet = workbook.active

    # Write the data to the sheet
    for row in data:
        sheet.append(row)

    # Set the width of the second column
    sheet.column_dimensions['A'].width = 30
    sheet.column_dimensions['B'].width = 45

    c1 = 'A1'
    c2 = 'A2'
    c3 = 'A3'
    c4 = 'A4'
    c5 = 'A5'
    c6 = 'A6'
    c7 = 'A7'
    c8 = 'A8'
    c9 = 'A9'
    c10 = 'A10'

    # Replace with the desired cell reference
    row_height =20  # Replace with the desired row height value

    # Set the row height for the specified cell
    sheet.row_dimensions[int(c1[1:])].height = row_height
    sheet.row_dimensions[int(c2[1:])].height = row_height
    sheet.row_dimensions[int(c3[1:])].height = row_height
    sheet.row_dimensions[int(c4[1:])].height = row_height
    sheet.row_dimensions[int(c5[1:])].height = row_height
    sheet.row_dimensions[int(c6[1:])].height = row_height
    sheet.row_dimensions[int(c7[1:])].height = row_height
    sheet.row_dimensions[int(c8[1:])].height = row_height
    sheet.row_dimensions[int(c9[1:])].height = row_height
    sheet.row_dimensions[int(c10[1:])].height = row_height

    # Align the text to the left in the second column for better visibility
    for cell in sheet['B']:
        cell.alignment = Alignment(horizontal='center')

    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    for row in sheet.iter_rows():
        for cell in row:
            cell.border = border

    filename = 'customer_data.xlsx'
    bold_cells = ['A1','A2','A3','A4','A5','A6','A7','A8','A9','A10', 'B1','B2','B3','B4','B5','B6','B7','B8','B9','B10']

    # Apply bold formatting to specified cells
    bold_font = Font(bold=True)
    for cell_ref in bold_cells:
        sheet[cell_ref].font = bold_font

    filename = 'customer_data.xlsx'
    image_filename = 'app_logo2.png'

    # Load the image
    img = Image(image_filename)

    # Adjust the size of the image as per your requirement
    img.width = 250
    img.height = 50

    # Add the image to a specific cell
    sheet.add_image(img, 'A1')  # Replace 'D5' with the desired cell reference

    # Save the workbook
    workbook.save(filename)

    print(f"Excel file '{filename}' has been created successfully!")

    # Get the printer name
    printer_name = win32print.GetDefaultPrinter()
    
    # Print the sheet
    win32api.ShellExecute(0, "print", filename, f'"{printer_name}"', ".", 0)



def reset():
    listbox1.delete(0,END)
    name_var.set("")
    ac_int.set("")
    bankai.set("")
    amt_int.set("")
    

def Rec(event):
    try:
        global selected_tuple
        index=listbox1.curselection()[0]
        selected_tuple=listbox1.get(index)

        ac_int.set(selected_tuple[0])
        bankai.set(selected_tuple[2])
        name_var.set(selected_tuple[3])

    except IndexError:
            pass
    

def filter_options(event):
    typed_text = lbox1.get().lower()
    matching_options = [item for item in items if item.lower().startswith(typed_text)]
    lbox1['values'] = matching_options
    
root.bind('<Key>', filter_options)

def convertTuplea(tup):
    st = '\n'.join(map(str, tup))
    return st

def submit():
    name=name_var.get()
    ac=ac_int.get()
    bank=bankai.get()
    amount=int(amt_int.get())
    mob=mob_int.get()
    mobl=len(mob_int.get())
    
    if name=="" or ac=="" or bank=="" or mob=="":
        print("empty")
    elif mobl!=10:
        print("mob no. invalid")
    else:
        today = date.today()
        date1=today.strftime("%Y-%m-%d")
        mycon=mys.connect(host="localhost",username="root",passwd='sshy6t5@',database="thoufeeq")
        cursor=mycon.cursor()
        cursor.execute("insert into sbi_csp_register(mob_no,ac_no,bank_name,name,amount,trdate) values('{}','{}','{}','{}',{},'{}');".format(mob,ac,bank,name,amount,date1))
        mycon.commit()
        mycon.close()

def search():
    mob=mob_int.get()
    mycon=mys.connect(host="localhost",username="root",passwd='sshy6t5@',database="thoufeeq")
    cursor=mycon.cursor()
    cursor.execute("select DISTINCT ac_no,mob_no,bank_name,name from sbi_csp_register where mob_no='{}'".format(mob,))
    dat=cursor.fetchall()
    rec=[]
    """for j in dat:
        for k in j:
            rec.append(k)
    print(rec)"""
    mycon.close()
    count=0
    for i in dat:
        count+=1
        a=convertTuplea(i)
        count1=str(count)
        c=convertTuplea(count1)
        listbox1.insert(END,i)
            
#------------------------------------------------LISTBOX---------------------------------------------------
listbox1=Listbox(root,font=('Arial',13,'bold'),width=50,height=11)
listbox1.bind('<<ListboxSelect>>',Rec)
listbox1.place(x="700",y="150")
sb_2.config(command=listbox1.yview)
#------------------------------------------------BUTTONS---------------------------------------------------
b1=Button(root,text="SUBMIT",font=("Arial",8,'bold'),width="10",height="2",activebackground="cyan",relief="ridge",command=submit)
b1.place(x="400",y="400")

b2=Button(root,text="🔎",font=("Arial",12,'bold'),width="5",activebackground="cyan",relief="ridge",command=search)
b2.place(x="625",y="150")

b3=Button(root,text="RESET",font=("Arial",8,'bold'),width="10",height="2",activebackground="cyan",relief="ridge",command=reset)
b3.place(x="1075",y="400")

b4=Button(root,text="PRINT",font=("Arial",8,'bold'),width="10",height="2",activebackground="cyan",relief="ridge",command=print1)
b4.place(x="75",y="400")

b6=Button(root,text="DENOMINATION",font=("Arial",12,'bold'),activebackground="cyan",relief="ridge",command=open1)
b6.place(x="1200",y="10")
#------------------------------------------------**END**---------------------------------------------------
root.mainloop()
