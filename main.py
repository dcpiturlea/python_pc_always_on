import win32com.client
import time
import tkinter
from multiprocessing import Process
import model
import datetime
from tkinter import ttk
from tkinter import messagebox
import os
import platform

def count_down(hour, minute, second):
    try:
        # the input provided by the user is
        # stored in here :temp
        temp = hour * 3600 + minute * 60 + second
    except:
        print("Please input the right value")
    while temp > -1 and stopped == False:

        # divmod(firstvalue = temp//60, secondvalue = temp%60)
        mins, secs = divmod(temp, 60)

        # Converting the input entered in mins or secs to hours,
        # mins ,secs(input = 110 min --> 120*60 = 6600 => 1hr :
        # 50min: 0sec)
        hours = 0
        if mins > 60:
            # divmod(firstvalue = temp//60, secondvalue
            # = temp%60)
            hours, mins = divmod(mins, 60)

        # using format () method to store the value up to
        # two decimal places
        lbl_hour['text'] = hours
        lbl_minute['text'] = mins
        lbl_second['text'] = secs
        # updating the GUI window after decrementing the
        # temp value every time
        window.update()
        time.sleep(1)

        # when temp value = 0; then a messagebox pop's up
        # with a message:"Time's up"
        if (temp == 0 and stopped == False):
            #window.after(10, window.destroy)
            #Message(title="Shut Down", message="Your PC will be turned off", master=window).show()
            shut_down_console()

        # after every one sec the value of temp will be decremented
        # by one
        temp -= 1

def shut_down_console():
    #strShutdown = "shutdown.exe -s -t " + str(0) + " -f"
    if platform.system() == "Windows":
        os.system("shutdown -s -t 0")

    else:
        os.system("shutdown -h now")

def btn_shut_down():
    global stopped
    stopped = False
    if btn_sd['text'] == 'Shut Down PC':
        btn_sd['text'] = 'Stop'
        btn_sd['bg'] = 'blue'
        lbl_sd['text'] = 'Status: Running'


        global window2
        window2 = tkinter.Tk()
        window2.title("Always On PC V1.0")
        window2.resizable(True, True)
        window2.wm_iconbitmap('iotech_logo.ico')
        window2.config(bg='white')

        lbl2 = tkinter.Label(window2, text="Hour:", width=15, bg='white')
        lbl2.config(font=('Arial 12 bold'))
        lbl2.grid(column=0, row=0)
        left_hours = model.get_hours_left()
        left_hours.sort()

        n = tkinter.StringVar()
        global hour_choosen
        hour_choosen = ttk.Combobox(window2, width=10, textvariable=n)
        hour_choosen['values'] = left_hours
        hour_choosen.current(0)
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

        now = datetime.datetime.now()
        now_min = now.minute

        now_min = int(now_min)
        new_left_min_left = []
        for i in range(1, len(left_mins) + 1):
            if i >= now_min:
                new_left_min_left.append(i)

        global mins_choosen
        mins_choosen = ttk.Combobox(window2, width=10, textvariable=n)
        mins_choosen['values'] = left_mins
        mins_choosen.current(now_min + 1)
        mins_choosen.grid(column=1, row=1)

        lbl5_bt = tkinter.Label(window2, text="", width=5, bg='white')
        lbl5_bt.grid(column=3, row=0)
        lbl6_bt = tkinter.Label(window2, text="", width=5, bg='white')
        lbl6_bt.grid(column=3, row=2)

        btn_sd2 = tkinter.Button(window2, text="Set", command=shut_down_pc, width=20, height=2, bg='gray', fg='white')
        btn_sd2.grid(column=3, row=1)

        window2.mainloop()
    elif btn_sd['text'] == 'Stop':
        btn_sd['text'] = 'Shut Down PC'
        lbl_sd['text'] = 'Status: Stopped'
        btn_sd['bg'] = 'gray'
        stopped = True
        lbl_hour['text'] = ''
        lbl_minute['text'] = ''
        lbl_second['text'] = ''



def shut_down_pc():
    btn_sd['text'] = 'Stop'
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
    else:
        #btn_sd['state'] = DISABLED
        window2.destroy()
        count_down(total_h, total_min, 0)



def btn__ao_click():
    global p
    if btn_ao['text'] == 'Always ON':
        p = Process(target=always_on_pc, args=())
        p.start()
        btn_ao['text'] = 'Stop'
        lbl['text'] = 'Status: Running'
        btn_ao['bg'] = 'red'
    elif btn_ao['text'] == 'Stop':
        p.terminate()
        btn_ao['text'] = 'Always ON'
        lbl['text'] = 'Status: Stopped'
        btn_ao['bg'] = 'gray'


def always_on_pc():
    shell = win32com.client.Dispatch("WScript.Shell")
    shell.AppActivate("Command Prompt")
    while True:
        shell.SendKeys("{NUMLOCK}")
        print("Numlock on")

        shell.SendKeys("{NUMLOCK}")
        print("Numlock Of")
        time.sleep(57)


def close_app():
    print("Close app")
    try:
        p.terminate()
        print("Process terminated")
    except:
        "An error has occured"
    finally:
        window.destroy()
        exit()


if __name__ == "__main__":
    window = tkinter.Tk()

    # Code to add widgets will go here...
    window.title("Always On PC V1.0")

    window.resizable(False, False)
    window.wm_iconbitmap('iotech_logo.ico')
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

    lbl_count_down = tkinter.Label(window, text="", width=2, bg='white')
    lbl_count_down.grid(column=2, row=6)

    lbl_hour = tkinter.Label(lbl_count_down, text="", width=2, bg='white', fg='red')
    lbl_hour.grid(column=2, row=6)


    lbl_minute = tkinter.Label(lbl_count_down, text="", width=2, bg='white', fg='red')
    lbl_minute.grid(column=3, row=6)


    lbl_second = tkinter.Label(lbl_count_down, text="", width=2, bg='white', fg='red')
    lbl_second.grid(column=4, row=6)

    window.protocol("WM_DELETE_WINDOW", close_app)


    window.mainloop()


