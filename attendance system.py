import winsound
import os
import time
import copy
from pyzbar.pyzbar import decode
import tkinter as tk
from PIL import Image, ImageTk
from datetime import datetime, timedelta
import pandas 
import cv2
import icecream as ic
from tkinter import messagebox
from tkinter import ttk

def get_tdy_date():
    return time.strftime('%Y_%m_%d', time.localtime())
#########################################################    change_time_state ### 
def change_time_state():
    global time_state, student_list
    if time_state == 0:
        time_state = 1

    elif time_state == 1:
        time_state = 0

    #########################################################    change_time_state ### gui edit
    if time_state == 0:
        # time in gui edit
        label5.configure(text="time in",)
        frame1.configure(bg='blue')
        label1.configure(bg='blue')

    elif time_state == 1:
        # time out gui edit
        label5.configure(text="time out",)
        frame1.configure(bg='green')
        label1.configure(bg='green')

    #########################################################    change_time_state ### update_list
    # after update gui need update list
    update_list(student_list)



#########################################################    check_date_data ### 
def check_date_data(student_list):
    # get now year and now month time date data
    now_year = time.strftime('%Y', time.localtime())
    now_month = time.strftime('%#m', time.localtime())
    
    #########################################################    check_date_data ### try get last date data
    try:
        f = open("student_data/last_date.txt", "r")
        now_ans = f.read()

        now_data = now_ans.split("_")
        last_time = datetime(int(now_data[0]), int(now_data[1]), int(now_data[2]))

        f.close()
        print(now_year)#2024
        print(last_time.year)#2022
        print(now_month)#8
        print(last_time.month)#7
        if now_year > str(last_time.year):
            # change now_year to 0
            # and now_month to 0
            
            for x in student_list:
                student_list[x]["attendance_days"] = 0
                student_list[x]["attendance_by_month"] = 0
                
        elif now_month != str(last_time.month):
            # change now_month to 0
            for x in student_list:
                student_list[x]["attendance_by_month"] = 0
        
    except FileNotFoundError as e:
        # run log file, afraid the file is too large, need many time loading, just pass first
        # dont change any thing
        messagebox.showerror("Error", f"File not found: {e}")
    except PermissionError as e:
        messagebox.showerror("Error", f"Permission denied: {e}")
    except Exception as e:
        messagebox.showerror("Error", f"An unexpected error occurred: {e}")
    

    #########################################################    check_date_data ###     return student_list
    return student_list
    
    

#########################################################    get_student_list ### 
def get_student_list():
    try:
        
        date_now = get_tdy_date()
        date_now_list = date_now + ".xlsx"

        day_ans = os.popen("dir /b day_data")
        day_list = day_ans.read().split("\n")
        print(day_list)

        #########################################################    get_student_list ### reload data
        # check today already open or not, if alr open reload the data
        if date_now_list in day_list:
            student_data = pandas.read_excel("day_data/" + date_now + ".xlsx") 
            student_list = student_data.to_dict('index')
            
            for x in student_list:
                if student_list[x]["date"] == 0:
                    student_list[x]["date"] = "0"
                if student_list[x]["time_in"] == 0:
                    student_list[x]["time_in"] = "0"
                if student_list[x]["time_out"] == 0:
                    student_list[x]["time_out"] = "0"
                if student_list[x]["attendance_days"] == 0:
                    pass
                if student_list[x]["attendance_by_month"] == 0:
                    pass
                if student_list[x]["attendance_rate"] == 0:
                    student_list[x]["attendance_rate"] = "0"
                if student_list[x]["attendance_rate_by_month"] == 0:
                    student_list[x]["attendance_rate_by_month"] = "0"
        
        #########################################################    get_student_list ### create data
        # if today is first time open           
        else:
            
            student_data = pandas.read_excel(r"CCM MASTER DATABASE_UPDATED YR 2024.xlsx",sheet_name='DataBase (SAT)')

                
            student_list = student_data.to_dict('index')

            for x in student_list:
                student_list[x]["date"] = "0"
                student_list[x]["time_in"] = "0"
                student_list[x]["time_out"] = "0"

                if str(student_list[x]["attendance_days"]) == 'nan':
                    student_list[x]["attendance_days"] = 0
                if str(student_list[x]["attendance_by_month"]) == 'nan':
                    student_list[x]["attendance_by_month"] = 0

                if student_list[x]["attendance_rate"] == 0 or str(student_list[x]["attendance_rate"]) == 'nan':
                    student_list[x]["attendance_rate"] = "0"
                if student_list[x]["attendance_rate_by_month"] == 0  or str(student_list[x]["attendance_rate_by_month"]) == 'nan':
                    student_list[x]["attendance_rate_by_month"] = "0"

        for x in student_list:
            print(student_list[x])
    except FileNotFoundError as e:
        messagebox.showerror("Error", f"File not found: {e}")
    except PermissionError as e:
        messagebox.showerror("Error", f"Permission denied: {e}")
    except Exception as e:
        messagebox.showerror("Error", f"An unexpected error occurred: {e}")
        #########################################################    get_student_list ###      return student_list
    return student_list



#########################################################    decode_func ###
# decode QR code or bar code
def decode_func():
    success, img = cap.read()
    mydata = 0
    for barcode in decode(img):
        mydata = barcode.data.decode('utf-8')
        print(mydata)
        time.sleep(1)
    return mydata
    


#########################################################    count_saturday ###
# count this year and this month have how many saturday
def count_saturday(ans):
    time_data = ans.split("/")
    join_time = datetime(int(time_data[0]), int(time_data[1]), int(time_data[2]))
    year_ans = time.strftime('%Y', time.localtime())
    # a = this year data
    # if user join time is after then a (this year data) , count from join time
    a = datetime(int(year_ans), 1, 1)
    if a < join_time:
        a = join_time
    this_month = 0
    this_year = 0

    #########################################################    count_saturday ###   start count
    while True:
        if a == datetime(int(year_ans)+1, 1, 1):
            break

        #########################################################    count_saturday ###   count this year have how many saturday
        if a.weekday() == 5:
            this_year +=1
            #print(a)
        
        #########################################################    count_saturday ###   count this month have how many saturday    
        month_ans = str(a.month)
        currnt_month_ans = time.strftime('%#m', time.localtime())
        if month_ans == currnt_month_ans:
            if a.weekday() == 5:
                this_month +=1
                #print(a)

        # b == one day
        # a + one day
        b = timedelta(1)
        a = a + b

    #########################################################    count_saturday ###     return this_month, this_year  
    return this_month, this_year



#########################################################    attendance ###    
def attendance(student_list, mydata):

    global time_state

    #########################################################    attendance ### time in
    if time_state == 0:
        ### time in
        for x in student_list:
            if mydata == student_list[x]["编号"] and student_list[x]["time_in"] == "0":
                datenow = get_tdy_date()
                student_list[x]["date"] = datenow
                timenow = time.strftime('%I:%M:%S %p', time.localtime())
                student_list[x]["time_in"] = timenow

                student_list[x]["attendance_days"] += 1
                student_list[x]["attendance_by_month"] += 1

                #########################################################    attendance ### count_saturday
                student_list[x]["生日日期"] = student_list[x]["生日日期"].strftime('%Y/%m/%d')
                this_month, this_year = count_saturday(student_list[x]["生日日期"])
        

                student_list[x]["attendance_rate"] = str(student_list[x]["attendance_days"]) + '/' + str(this_year)
                student_list[x]["attendance_rate_by_month"] = str(student_list[x]["attendance_by_month"]) + '/' + str(this_month)
                
                print_out(student_list)
          
        # sound output
        duration = 300  # millisecond
        freq = 650  # Hz
        winsound.Beep(freq, duration)

        #########################################################    attendance ###     return student_list
        return student_list


    #########################################################    attendance ### time out
    elif time_state == 1:
        ### time out
        
        
        
        for x in student_list:
            if mydata  == student_list[x]["编号"]:
                if student_list[x]["time_in"] == "0":
                    print("Error: Student has no 'time in' data.")
                    messagebox.showerror("Error", "Student has no 'time in' data.")
            elif mydata == student_list[x]["编号"] and student_list[x]["time_out"] == "0":
                datenow = time.strftime('%Y-%m-%d', time.localtime())
                student_list[x]["date"] = datenow
                timenow = time.strftime('%I:%M:%S %p', time.localtime())
                student_list[x]["time_out"] = timenow

                print_out(student_list)
            '''if mydata == student_list[x]["编号"] and student_list[x]["time_out"] == "0":
                datenow = time.strftime('%Y-%m-%d', time.localtime())
                student_list[x]["date"] = datenow
                timenow = time.strftime('%I:%M:%S %p', time.localtime())
                student_list[x]["time_out"] = timenow
            
                print_out(student_list)
                '''

        # sound output
        duration = 300  # millisecond
        freq = 650  # Hz
        winsound.Beep(freq, duration)

        #########################################################    attendance ###     return student_list
        return student_list

        


#########################################################    update_list ###
def update_list(student_list):
    displaylistin = []
    displaylistout = []
    global time_state

    #########################################################    update_list ### time in
    if time_state == 0:

        treeview.delete(*treeview.get_children())
        # get have "time_in" data user
        print_list = {}
        for x in student_list:
            if student_list[x]["time_in"] != "0":
                print_data = student_list[x].copy()
            
                print_list[x] = print_data.copy()
                
        
        # if no user have "time_in" data 
        if len(print_list) == 0:
            print("timein")
            ic.ic(student_list)
            '''
            textlist1 = "name".center(25," ") + "\n" 
            textlist2 = "|" + "time_in".center(18," ")  + "\n"
            label3.configure(text=textlist1,)
            label2.configure(text=textlist2,)
            #########################################################    update_list ###      end
            return 0
            '''

        #########################################################    update_list ### arrange list
        # arrange print_list order to order_list 
        print(print_list)
        order_list = sorted(print_list, key=lambda x:student_list[x]["time_in"])
        order_list.reverse()
        
        # print print_list by order_list
        print_final_list = {}
        for x in order_list:
            if x in print_list:
                order_data = print_list[x].copy()
                
                print_final_list[x] = order_data.copy()

        # edit cmd and gui
        textlist1 = "name".center(25," ") + "\n" 
        textlist2 = "|" + "time_in".center(18," ")  + "\n"
        #os.system("cls")
        print("          name           " + "|     time_in         ")
        num_student = str(0) + "/" + str(len(student_list))
        y = 1
        for x in print_final_list:
            if print_final_list[x]["time_in"] != "0":
                if y <= 18:
                    
                    
                    #textlist1 += print_final_list[x]["孩子姓名(中)"].center(25," ") + "\n" 
                    #textlist2 += "|"+ print_final_list[x]["time_in"].center(18," ")  + "\n"
                    #label3.configure(text=textlist1,)
                    #label2.configure(text=textlist2,)
                    
                    entry = (print_final_list[x]["孩子姓名(中)"], print_final_list[x]["time_in"])
                    if entry not in displaylistin:
                        displaylistin.append(entry)
                        treeview.insert('', 'end', values=entry)

                '''num_student = str(len(print_final_list)) + "/" + str(len(student_list))
                label4.configure(text=num_student,)
                print(print_final_list[x]["孩子姓名(中)"].center(25," "),end="|")
                print(print_final_list[x]["time_in"].center(18," "))
                y += 1'''

        print(" ")
        print(num_student.rjust(43," "))
    
    #########################################################    update_list ### time out
    elif time_state == 1:
        
        treeview.delete(*treeview.get_children())
        # get have "time_out" data user
        print_list = {}
        for x in student_list:
            if student_list[x]["time_out"] != "0":
                print_data = student_list[x].copy()
            
                print_list[x] = print_data.copy()


        # if no user have "time_in" data 
        if len(print_list) == 0:

            print("timeout")
            ic.ic(student_list)
            '''
            textlist1 = "name".center(25," ") + "\n" 
            textlist2 = "|" + "time_out".center(18," ")  + "\n"
            label3.configure(text=textlist1,)
            label2.configure(text=textlist2,)
            #########################################################    update_list ###      end
            return 0'''
            

        #########################################################    update_list ### arrange list
        # arrange print_list order to order_list 
        print(print_list)
        order_list = sorted(print_list, key=lambda x:student_list[x]["time_out"])
        order_list.reverse()
        
        # print print_list by order_list
        print_final_list = {}
        for x in order_list:
            if x in print_list:
                order_data = print_list[x].copy()
                
                print_final_list[x] = order_data.copy()


        # edit cmd and gui
        textlist1 = "name".center(25," ") + "\n" 
        textlist2 = "|" + "time_out".center(18," ")  + "\n"
        #os.system("cls")
        print("          name           " + "|     time_out         ")
        num_student = str(0) + "/" + str(len(student_list))
        y = 1
        for x in print_final_list:
            if print_final_list[x]["time_out"] != "0":
                if y <= 18:
                    
                    entry = (print_final_list[x]["孩子姓名(中)"], print_final_list[x]["time_out"])
                    if entry not in displaylistout:
                        displaylistout.append(entry)
                        treeview.insert('', 'end', values=entry)                    
                    '''textlist1 += print_final_list[x]["孩子姓名(中)"].center(25," ") + "\n" 
                    textlist2 += "|"+ print_final_list[x]["time_out"].center(18," ")  + "\n"
                    label3.configure(text=textlist1,)
                    label2.configure(text=textlist2,)'''

                '''num_student = str(len(print_final_list)) + "/" + str(len(student_list))
                label4.configure(text=num_student,)
                print(print_final_list[x]["孩子姓名(中)"].center(25," "),end="|")
                print(print_final_list[x]["time_out"].center(18," "))
                y += 1'''

        print(" ")
        print(num_student.rjust(43," "))


#########################################################    print_out ### 
# print out t
# o excel
def print_out(student_list):
    
    #########################################################    print_out ### change student_list to pandas Full_Data_Frame
    Full_Data_Frame = pandas.DataFrame(student_list,)
    Full_Data_Frame = Full_Data_Frame.T
    print(Full_Data_Frame)
    
    #########################################################    print_out ### output day data
    date_now = get_tdy_date()
    Full_Data_Frame.to_excel("day_data/" + date_now + ".xlsx", index = False) 
    time_now = time.strftime('%Y_%m_%d', time.localtime()) + time.strftime(' %I_%M_%S %p', time.localtime())
    Full_Data_Frame.to_excel("day_data/backup/" + time_now + ".xlsx", index = False) 



    #########################################################    print_out ### edit original_data
    # xxx original_data = student_list cannot be use 
    # because it is list and it just copy position, need use deep copy
    original_data = copy.deepcopy(student_list)
    # for original_data dont have "date", "time_in", "time_out"
    for x in original_data:
        del original_data[x]["date"]
        del original_data[x]["time_in"]
        del original_data[x]["time_out"]
    #########################################################    print_out ### change original_data to pandas Full_Data_Frame
    Full_Data_Frame = pandas.DataFrame(original_data,)
    Full_Data_Frame = Full_Data_Frame.T
    #########################################################    print_out ### output student data
    Full_Data_Frame.to_excel("student_data/student_data.xlsx", index = False) 
    


#########################################################    add_month_data ###
def add_month_data(student_list):

    #########################################################    add_month_data ### get month data
    # check have current month data or not
    month_time = time.strftime("%B %Y") 
    month_time_xlsx = time.strftime("%B %Y") + ".xlsx"
    month_ans = os.popen("dir /b month_data")
    month_list = month_ans.read().split("\n")

    # if have month data, get month data
    if month_time_xlsx in month_list:
        output_list = pandas.read_excel("month_data/" + month_time + ".xlsx")
        
        for x in student_list:
            if student_list[x]["编号"] not in output_list["编号"].tolist():
                output_data = pandas.DataFrame([[student_list[x]["编号"],student_list[x]["english_name"]]],
                                columns = ["编号","name"])
                output_list = pandas.concat([output_list,output_data])

        print(output_list)

    # if dont have month data, create month data
    else:
        output_list = pandas.DataFrame(columns = ["编号","孩子姓名(中)"])

        for x in student_list:
            output_data = pandas.DataFrame([[student_list[x]["编号"],student_list[x]["孩子姓名(中)"]]],
                        columns = ["编号","name"])
            output_list = pandas.concat([output_list,output_data])

    #########################################################    add_month_data ### count attendance data
    len_data = len(output_list.columns)
    time_date = time.strftime('%d')
    attendance_data = []
    for x in student_list:
        if student_list[x]["date"] != '0':
            attendance_data.append(1)
        else:
            attendance_data.append(0)
            
    #########################################################    add_month_data ### add attendance data
    output_list.insert(len_data, time_date, attendance_data,allow_duplicates=True)

    #########################################################    add_month_data ### output month data
    output_list.to_excel("month_data/" + month_time + ".xlsx", index = False) 

    
def change_date_data():
    f = open("student_data/last_date.txt", "w")
    date_now =  get_tdy_date()
    f.write(date_now)
    f.close()



#########################################################    main ###

student_list = get_student_list()

student_list = check_date_data(student_list)
ic.ic(student_list)

cap = cv2.VideoCapture(0)
time_state = 0

### graphic ###
root = tk.Tk()
#width, height = 1200, 650
root.title("CCM Attendance System")
root.geometry("1350x750")


# Configure grid
root.grid_columnconfigure(0, weight=3)
root.grid_columnconfigure(1, weight=2)
root.grid_rowconfigure(0, weight=3)
root.grid_rowconfigure(1, weight=2)


# Main frame (frame1)
frame1 = tk.Frame(master=root, bg='blue')
frame1.grid(row=1, column=0, sticky="nsew")

label1 = tk.Label(frame1, text="", font=('Arial', 48), bg='blue', fg='white')
label1.pack(expand=True, fill='both')


# Treeview style
style = ttk.Style()
style.configure("Treeview", font=("Helvetica", 18, "bold"), rowheight=40)
style.configure("Treeview.Heading", font=("Helvetica", 24, "bold"))

# Treeview frame
treeview_frame = tk.Frame(root)
treeview_frame.grid(row=0, column=1, rowspan=1, sticky="nsew")

columns = ('Name', 'Time')
treeview = ttk.Treeview(treeview_frame, columns=columns, show='headings', height=8)
treeview.heading('Name', text='名字')
treeview.heading('Time', text='时间')
treeview.column('Name', anchor=tk.CENTER , width=100)
treeview.column('Time', anchor=tk.CENTER, width=100)
treeview.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

scrollbar = ttk.Scrollbar(treeview_frame, orient=tk.VERTICAL, command=treeview.yview)
treeview.configure(yscroll=scrollbar.set)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

# Label frame
label_frame = tk.Frame(root)
label_frame.grid(row=1, column=1, columnspan=1, sticky="ew", padx=10, pady=10)

button6 = tk.Button(label_frame, text="change", font=('Arial', 18), borderwidth=8, command=change_time_state)
button6.pack(side=tk.LEFT, padx=5)

label5 = tk.Label(label_frame, text="time in", font=('Arial', 32), borderwidth=10)
label5.pack(side=tk.LEFT, padx=5)


# Main label for image display
lmain = tk.Label(root)
lmain.grid(row=0, column=0, sticky="nsew")


#update_list(student_list)

########################################################    main ### main loop
def main():
    global student_list
    
    #########################################################    main ### change camera display
    success, img = cap.read()

    frame = cv2.flip(img, 1)
    cv2image = cv2.cvtColor(frame, cv2.COLOR_BGR2RGBA)
    img = Image.fromarray(cv2image)
    x, y = (img.size)
    zoom_size = 1.4
    img = img.resize((int(x*zoom_size), int(y*zoom_size)))
    imgtk = ImageTk.PhotoImage(image=img)
    lmain.imgtk = imgtk
    lmain.configure(image=imgtk)
    
    #########################################################    main ###  change time
    timenow = time.strftime('%I:%M:%S %p', time.localtime())
    label1.configure(text=timenow)


    #########################################################    main ### decode fuction
    mydata = 0
    mydata = decode_func()
    if mydata != 0:
        
        if mydata not in [student['编号'] for student in student_list.values()]:
            print("Not in the list")
            messagebox.showerror("Error", "Student not in database")
        else:
            student_list = attendance(student_list, mydata)
            update_list(student_list)
            mydata = 0


    #########################################################    main ### Loop main function
    lmain.after(10, main)
    

main()

root.mainloop()

# if main loop end, run output data function
print_out(student_list)
add_month_data(student_list)
change_date_data()
