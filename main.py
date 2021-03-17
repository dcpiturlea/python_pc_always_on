import win32com.client
import time
import tkinter
from multiprocessing import Process

def btn_click():
    global p
    if btn['text'] == 'Always ON':
        p = Process(target=always_on_pc, args=())
        p.start()
        btn['text'] = 'Always OF'
        lbl['text'] = 'Running'
        btn['bg'] = 'red'
    elif btn['text'] == 'Always OF':
        p.terminate()
        btn['text'] = 'Always ON'
        lbl['text'] = 'Stopped'
        btn['bg'] = 'gray'


def always_on_pc():
    while True:
        shell = win32com.client.Dispatch("WScript.Shell")
        shell.AppActivate("Command Prompt")
        shell.SendKeys("{NUMLOCK}")
        print("Numlock on")
        time.sleep(1)
        shell.SendKeys("{NUMLOCK}")
        print("Numlock Of")
        time.sleep(180)


if __name__ == "__main__":
    window = tkinter.Tk()

    # Code to add widgets will go here...
    window.title("Always On PC V1.0")
    # window.geometry('600x200')
    window.resizable(False, False)

    window.config(bg='white')

    lbl = tkinter.Label(window, text="Stopped",width=30,bg='white')
    lbl.config(font=('Arial 12 bold'))
    lbl.grid(column=1, row=0)
    lblLoc = tkinter.Label(window, text="", bg='white')
    btn = tkinter.Button(window, text="Always ON", command=btn_click, width=20, height=2, bg='gray', fg='white')

    btn.grid(column=1, row=1)
    lblLoc.grid(column=1, row=4)
    lblTeam = tkinter.Label(window, text="ioTech.ro", bg='white')
    lblTeam.grid(column=1, row=6)
    window.mainloop()