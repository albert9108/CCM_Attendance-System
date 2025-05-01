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
        status_indicator.config(text="签到", bg="blue", fg="white")

    elif time_state == 1:
        # time out gui edit
        label5.configure(text="time out",)
        frame1.configure(bg='green')
        label1.configure(bg='green')
        status_indicator.config(text="签退", bg="green", fg="white")

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
                
                # Show student info in status bar
                status_text.config(text=f"上次扫描: {student_list[x]['孩子姓名(中)']} - 签到成功")
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
                    status_text.config(text=f"错误: 学生未签到")
                elif mydata == student_list[x]["编号"] and student_list[x]["time_out"] == "0":
                    datenow = time.strftime('%Y-%m-%d', time.localtime())
                    student_list[x]["date"] = datenow
                    timenow = time.strftime('%I:%M:%S %p', time.localtime())
                    student_list[x]["time_out"] = timenow
                    status_text.config(text=f"上次扫描: {student_list[x]['孩子姓名(中)']} - 签退成功")
                
                print("Here")
                print_out(student_list)

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
    global time_state, student_count_label

    # Clear current treeview
    treeview.delete(*treeview.get_children())

    #########################################################    update_list ### time in
    if time_state == 0:
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
            student_count_label.config(text="0/" + str(len(student_list)))
            return

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

        # Insert into treeview
        for x in print_final_list:
            if print_final_list[x]["time_in"] != "0":
                entry = (print_final_list[x]["孩子姓名(中)"], print_final_list[x]["time_in"])
                if entry not in displaylistin:
                    displaylistin.append(entry)
                    treeview.insert('', 'end', values=entry)
                    
        # Update student count
        student_count_label.config(text=f"{len(print_final_list)}/{len(student_list)}")
        print(" ")
    
    #########################################################    update_list ### time out
    elif time_state == 1:
        # get have "time_out" data user
        print_list = {}
        for x in student_list:
            if student_list[x]["time_out"] != "0":
                print_data = student_list[x].copy()
                print_list[x] = print_data.copy()

        # if no user have "time_out" data 
        if len(print_list) == 0:
            print("timeout")
            student_count_label.config(text="0/" + str(len(student_list)))
            return

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

        # Insert into treeview
        for x in print_final_list:
            if print_final_list[x]["time_out"] != "0":
                entry = (print_final_list[x]["孩子姓名(中)"], print_final_list[x]["time_out"])
                if entry not in displaylistout:
                    displaylistout.append(entry)
                    treeview.insert('', 'end', values=entry)
                    
        # Update student count
        student_count_label.config(text=f"{len(print_final_list)}/{len(student_list)}")
        print(" ")


#########################################################    print_out ### 
# print out to excel
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
                output_data = pandas.DataFrame([[student_list[x]["编号"],student_list[x]["孩子姓名(中)"]]],
                                columns = ["编号","孩子姓名(中)"])
                output_list = pandas.concat([output_list,output_data])

        print(output_list)

    # if dont have month data, create month data
    else:
        output_list = pandas.DataFrame(columns = ["编号","孩子姓名(中)"])

        for x in student_list:
            output_data = pandas.DataFrame([[student_list[x]["编号"],student_list[x]["孩子姓名(中)"]]],
                        columns = ["编号","孩子姓名(中)"])
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
    
    # Update or add the date column
    if time_date in output_list.columns:
        output_list[time_date] = attendance_data  # Override existing column
    else:
        output_list[time_date] = attendance_data  # Add new column

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
root.title("CCM Attendance System")
root.geometry("1350x750")
root.minsize(1000, 650)  # Set minimum window size

# Create a custom style
style = ttk.Style()
style.configure("Treeview", 
                font=("Helvetica", 18, "bold"), 
                rowheight=40)
style.configure("Treeview.Heading", 
                font=("Helvetica", 20, "bold"))
style.map("Treeview", 
          background=[("selected", "#3366CC")],
          foreground=[("selected", "white")])

# Configure grid with weights
root.grid_columnconfigure(0, weight=3)
root.grid_columnconfigure(1, weight=1)
root.grid_rowconfigure(0, weight=10)
root.grid_rowconfigure(1, weight=2)
root.grid_rowconfigure(2, weight=1)

# Main frame for camera display
camera_frame = tk.Frame(root)
camera_frame.grid(row=0, column=0, sticky="nsew")

# Clock frame (frame1)
frame1 = tk.Frame(master=root, bg='blue')
frame1.grid(row=1, column=0, sticky="nsew")

# Clock display
label1 = tk.Label(frame1, text="", font=('Arial', 48, 'bold'), bg='blue', fg='white')
label1.pack(expand=True, fill='both')

# Status bar frame
status_bar = tk.Frame(root, height=30, bg='#f0f0f0')
status_bar.grid(row=2, column=0, columnspan=2, sticky="ew")
status_bar.grid_columnconfigure(0, weight=1)
status_bar.grid_columnconfigure(1, weight=1)

# Status text (shows last scanned student)
status_text = tk.Label(status_bar, text="准备就绪", font=('Arial', 12), bg='#f0f0f0')
status_text.grid(row=0, column=0, sticky="w", padx=10)

# Current date display
date_label = tk.Label(status_bar, text=f"日期: {time.strftime('%Y-%m-%d')}", font=('Arial', 12), bg='#f0f0f0')
date_label.grid(row=0, column=1, sticky="e", padx=10)

# Right side panel - contains the treeview and controls
right_panel = tk.Frame(root)
right_panel.grid(row=0, column=1, rowspan=2, sticky="nsew")
right_panel.grid_columnconfigure(0, weight=1)
right_panel.grid_rowconfigure(0, weight=0)  # Header
right_panel.grid_rowconfigure(1, weight=1)  # Treeview
right_panel.grid_rowconfigure(2, weight=0)  # Controls

# Header for the right panel
header_frame = tk.Frame(right_panel, bg='#3366CC', height=50)
header_frame.grid(row=0, column=0, sticky="ew")

# Status indicator (Sign-in/Sign-out)
status_indicator = tk.Label(header_frame, text="签到", font=('Arial', 20, 'bold'), bg='blue', fg='white')
status_indicator.pack(side=tk.LEFT, padx=10, pady=5)

# Student count label
student_count_label = tk.Label(header_frame, text="0/0", font=('Arial', 16), bg='#3366CC', fg='white')
student_count_label.pack(side=tk.RIGHT, padx=10, pady=5)

# Treeview frame
treeview_frame = tk.Frame(right_panel)
treeview_frame.grid(row=1, column=0, sticky="nsew")
treeview_frame.grid_columnconfigure(0, weight=1)
treeview_frame.grid_rowconfigure(0, weight=1)

# Create the treeview with columns
columns = ('Name', 'Time')
treeview = ttk.Treeview(treeview_frame, columns=columns, show='headings')
treeview.heading('Name', text='名字')
treeview.heading('Time', text='时间')
treeview.column('Name', anchor=tk.CENTER, width=150)
treeview.column('Time', anchor=tk.CENTER, width=150)
treeview.grid(row=0, column=0, sticky="nsew")

# Add scrollbar to treeview
scrollbar = ttk.Scrollbar(treeview_frame, orient=tk.VERTICAL, command=treeview.yview)
scrollbar.grid(row=0, column=1, sticky="ns")
treeview.configure(yscrollcommand=scrollbar.set)

# Controls frame
controls_frame = tk.Frame(right_panel, bg='#f0f0f0', height=100)
controls_frame.grid(row=2, column=0, sticky="ew", padx=5, pady=5)

# Mode toggle button
mode_button = tk.Button(
    controls_frame, 
    text="切换模式", 
    font=('Arial', 16), 
    bg='#3366CC', 
    fg='white', 
    borderwidth=5, 
    command=change_time_state
)
mode_button.pack(side=tk.LEFT, padx=10, pady=10)

# Mode label
label5 = tk.Label(controls_frame, text="time in", font=('Arial', 16, 'bold'))
label5.pack(side=tk.LEFT, padx=10, pady=10)

# Main label for image display
lmain = tk.Label(camera_frame)
lmain.pack(expand=True, fill='both')

########################################################    main ### main loop
def main():
    global student_list
    
    #########################################################    main ### change camera display
    success, img = cap.read()
    
    if success and img is not None:
        frame = cv2.flip(img, 1)
        cv2image = cv2.cvtColor(frame, cv2.COLOR_BGR2RGBA)
        img = Image.fromarray(cv2image)
        
        # Get parent widget size and calculate zoom
        w = camera_frame.winfo_width()
        h = camera_frame.winfo_height()
        
        # if w <= 0 or h <= 0:
        #     # Use default dimensions until the frame is properly sized
        #     img = img.resize((640, 480))
        #     print("Using default dimensions until camera frame is sized")
        
        if w > 0 and h > 0:  # Ensure dimensions are valid
            img_w, img_h = img.size

            # Calculate the aspect ratio
            img_ratio = img_w / img_h
            frame_ratio = w / h

            if frame_ratio > img_ratio:
                # Height is the limiting factor
                new_h = h
                new_w = int(h * img_ratio)
            else:
                # Width is the limiting factor
                new_w = w
                new_h = int(w / img_ratio)
            if new_w > 0 and new_h > 0:
                img = img.resize((new_w, new_h))
            else:
               print(f"Warning: Invalid resize dimensions calculated: {new_w}x{new_h}") 
        else:
            print("Warning: Camera frame dimensions are not set yet.")

            
        imgtk = ImageTk.PhotoImage(image=img)
        lmain.imgtk = imgtk
        lmain.configure(image=imgtk)
    else:
        messagebox.showerror("Error", "Failed to capture image from camera")
    
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
            status_text.config(text="错误: 学生不在数据库中")
        else:
            student_list = attendance(student_list, mydata)
            update_list(student_list)
            mydata = 0


    #########################################################    main ### Loop main function
    lmain.after(10, main)
    

# Update list when first open
update_list(student_list)

# Start the main loop
main()

root.mainloop()

# if main loop end, run output data function
print_out(student_list)
add_month_data(student_list)
change_date_data()
