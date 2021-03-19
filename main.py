from tkinter.messagebox import Message
import threading
import win32com.client
import time
import tkinter
import model
import datetime
from tkinter import ttk
from tkinter import DISABLED, NORMAL
import os
import platform
from os import path

"""
impachetare: 
pyinstaller --onefile --noconsole main.py
pyinstaller --noconsole iotech_logo.ico main.py
pyinstaller --onefile --windowed --icon="shutdown.ico" --add-data=iotech_logo.ico;icon  --noconsole main.py
"""


def count_down(hour, minute, time_to_numlock):
    if hour < 10:
        lbl_hour['text'] = "0" + str(hour)
    else:
        lbl_hour['text'] = hour
    if minute < 10:
        lbl_minute['text'] = "0" + str(minute)
    else:
        lbl_minute['text'] = minute
    lbl_points['text'] = ':'
    window.update()
    if time_to_numlock > 0:
        time_to_numlock = 1000
    else:
        time_to_numlock = 72500
    new_time_to_numlock = time_to_numlock
    while stopped == False:
        window.update()
        time.sleep(0.05)
        if stopped == False and hour == datetime.datetime.now().hour and minute == datetime.datetime.now().minute:
            # window.after(10, window.destroy)
            #Message(title="Shut Down", message="Your PC will be turned off", master=window).show()
            lbl_hour['font'] = 'Helvetica 10 bold'
            lbl_minute['font'] = 'Helvetica 10 bold'
            lbl_points['font'] = 'Helvetica 10 bold'
            lbl_hour['width'] = 5
            lbl_minute['width'] = 5
            lbl_points['width'] = 5
            lbl_hour['text'] = "Shut"
            lbl_minute['text'] = "Now"
            lbl_points['text'] = "Down"
            window.update()
            time.sleep(5)
            btn_sd['state'] = DISABLED
            shut_down_console()
            break
        if time_to_numlock == 1:
            shell.SendKeys("{NUMLOCK}")
            print("Numlock on: sd")
            shell.SendKeys("{NUMLOCK}")
            print("Numlock Of: sd")
            time_to_numlock = new_time_to_numlock
        time_to_numlock = time_to_numlock - 1



def shut_down_console():
    # strShutdown = "shutdown.exe -s -t " + str(0) + " -f"
    if platform.system() == "Windows":
        os.system('shutdown /p /f')

    else:
        os.system("shutdown -h now")


def btn_shut_down():
    global stopped
    stopped = False
    if btn_sd['text'] == 'Shut Down PC':
        btn_sd['state'] = DISABLED
        global window2
        window2 = tkinter.Tk()
        window2.title("Always On PC V1.0")
        window2.resizable(True, True)
        window2.wm_iconbitmap(path.abspath(path.join(path.dirname(__file__), 'iotech_logo.ico')))
        window2.config(bg='white')

        lbl2 = tkinter.Label(window2, text="Hour:", width=15, bg='white')
        lbl2.config(font=('Arial 12 bold'))
        lbl2.grid(column=0, row=0)
        left_hours = model.get_hours_left()
        left_hours.sort()
        now = datetime.datetime.now()

        now_h = int(str(now.hour))
        n = tkinter.StringVar()
        global hour_choosen
        hour_choosen = ttk.Combobox(window2, width=10, textvariable=n)
        hour_choosen['values'] = left_hours
        hour_choosen.current(now_h -1)
        hour_choosen.grid(column=0, row=1)

        lbl4 = tkinter.Label(window2, text="", width=5, bg='white')
        lbl4.grid(column=2, row=1)
        lbl5 = tkinter.Label(window2, text="", width=5, bg='white')
        lbl5.grid(column=4, row=1)
        lbl6 = tkinter.Label(window2, text="", width=5, bg='white')
        lbl6.grid(column=4, row=3)
        lbl7 = tkinter.Label(window2, text="", width=2, bg='white')
        lbl7.grid(column=4, row=0)

        lbl3 = tkinter.Label(window2, text="Min:", width=15, bg='white')
        lbl3.config(font=('Arial 12 bold'))
        lbl3.grid(column=1, row=0)

        n = tkinter.StringVar()
        left_mins = model.get_all_min()


        now_min = now.minute

        now_min = int(now_min)
        new_left_min_left = []
        for i in range(1, len(left_mins) + 1):
            if i >= now_min:
                new_left_min_left.append(i)

        global mins_choosen
        mins_choosen = ttk.Combobox(window2, width=10, textvariable=n)
        mins_choosen['values'] = left_mins
        if now_min == 59:
            mins_choosen.current(now_min)
        else:
            mins_choosen.current(now_min + 1)
        mins_choosen.grid(column=1, row=1)

        lbl5_bt = tkinter.Label(window2, text="", width=5, bg='white')
        lbl5_bt.grid(column=3, row=0)
        lbl6_bt = tkinter.Label(window2, text="", width=5, bg='white')
        lbl6_bt.grid(column=3, row=2)

        btn_set_time = tkinter.Button(window2, text="Set", command=set_time_to_power_off, width=20, height=2, bg='gray',
                                      fg='white')
        btn_set_time.grid(column=3, row=1)

        window2.protocol("WM_DELETE_WINDOW", close_window_2)
        window2.mainloop()

    elif btn_sd['text'] == 'Stop Shut Down PC':
        btn_sd['text'] = 'Shut Down PC'
        lbl_sd['text'] = 'Status: Stopped'
        btn_sd['bg'] = 'gray'
        stopped = True
        lbl_hour['text'] = ''
        lbl_minute['text'] = ''
        lbl_points['text'] = ''


def set_time_to_power_off():
    btn_sd['text'] = 'Stop Shut Down PC'
    lbl_sd['text'] = 'Status: Running'
    btn_sd['bg'] = 'blue'
    now = datetime.datetime.now()
    now_h = now.hour
    now_m = now.minute
    t2_h = hour_choosen.get()
    t2_m = mins_choosen.get()
    t2_m = int(t2_m)
    t2_h = int(t2_h)

    total_mins = model.get_total_min_to_shut_down(int(now_h), int(now_m), t2_h, t2_m)
    total_sec = total_mins * 60
    total_h = int(total_mins / 60)
    total_min = total_mins % 60
    print("h " + str(total_h))
    print("m " + str(total_min))

    if total_mins < 1:
        tkinter.messagebox.showerror(title="Error", message="Please select a valide time")
        window2.destroy()
        btn_sd['bg'] = 'gray'
        btn_sd['text'] = 'Shut Down PC'
        btn_sd['state'] = NORMAL
    else:
        # btn_sd['state'] = DISABLED
        btn_sd['state'] = NORMAL
        window2.destroy()
        try:
            t2 = threading.Thread(target=count_down(t2_h, t2_m, time_to_numlock), args=())
            t2.run()
        except:
            close_app()
        #window.after(10, count_down(t2_h, t2_m))



def btn__ao_click():
    if btn_ao['text'] == 'Always ON':
        btn_ao['text'] = 'Stop Always ON'
        lbl['text'] = 'Status: Running'
        btn_ao['bg'] = 'red'
        t1 = threading.Thread(target=always_on_pc, args=())
        try:
            t1.run()
        except:
            close_app()
    elif btn_ao['text'] == 'Stop Always ON':
        btn_ao['text'] = 'Always ON'
        lbl['text'] = 'Status: Stopped'
        btn_ao['bg'] = 'gray'


def always_on_pc():
    shell.AppActivate("Command Prompt")
    global time_to_numlock
    time_to_numlock = 1000
    print("always on: on")
    while time_to_numlock <= 1000:
        if time_to_numlock == 1000:
            shell.SendKeys("{NUMLOCK}")
            print("Numlock on")
            shell.SendKeys("{NUMLOCK}")
            print("Numlock Of")
        window.update()
        time.sleep(0.05)
        time_to_numlock = time_to_numlock - 1
        if time_to_numlock == 1:
            time_to_numlock = 1000


def close_window_2():
    btn_sd['state'] = NORMAL
    window2.destroy()


def close_app():
    print("Close app")
    try:
        window.destroy()
        print("Process terminated")
    except:
        "An error has occured"
    finally:

        exit()


if __name__ == "__main__":
    window = tkinter.Tk()

    # Code to add widgets will go here...
    window.title("Always On PC V1.0")

    window.resizable(False, False)
    window.wm_iconbitmap(path.abspath(path.join(path.dirname(__file__), 'iotech_logo.ico')))
    window.config(bg='white')

    lbl = tkinter.Label(window, text="Status: Stopped", width=30, bg='white')
    lbl.config(font=('Arial 12 bold'))
    lbl.grid(column=1, row=0)
    lblLoc = tkinter.Label(window, text="", bg='white')
    btn_ao = tkinter.Button(window, text="Always ON", command=btn__ao_click, width=20, height=2, bg='gray', fg='white')
    btn_ao.grid(column=1, row=1)

    lblLoc.grid(column=1, row=4)
    lblTeam = tkinter.Label(window, text="ioTech.ro", bg='white')
    lblTeam.grid(column=1, row=6)

    lbl_sd = tkinter.Label(window, text="Status: Stopped", width=30, bg='white')
    lbl_sd.config(font=('Arial 12 bold'))
    lbl_sd.grid(column=2, row=0)
    lblLoc_sd = tkinter.Label(window, text="", bg='white')

    global btn_sd
    btn_sd = tkinter.Button(window, text="Shut Down PC", command=btn_shut_down, width=20, height=2, bg='gray',
                            fg='white')
    btn_sd.grid(column=2, row=1)

    lbl_count_down = tkinter.Label(window, text="", width=6, bg='white')
    lbl_count_down.grid(column=2, row=6)

    lbl_hour = tkinter.Label(lbl_count_down, text="", width=2, bg='white', fg='red', font='Helvetica 12 bold')
    lbl_hour.grid(column=3, row=6)

    lbl_minute = tkinter.Label(lbl_count_down, text="", width=2, bg='white', fg='red', font='Helvetica 12 bold')
    lbl_minute.grid(column=5, row=6)

    lbl_points = tkinter.Label(lbl_count_down, text="", width=2, bg='white', fg='red', font='Helvetica 12 bold')
    lbl_points.grid(column=4, row=6)
    global shell
    shell = win32com.client.Dispatch("WScript.Shell")
    global time_to_numlock
    time_to_numlock = 0
    window.protocol("WM_DELETE_WINDOW", close_app)

    window.mainloop()
