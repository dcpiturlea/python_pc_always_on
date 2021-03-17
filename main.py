import win32com.client
import time
import tkinter


def always_on_pc():
    shell = win32com.client.Dispatch("WScript.Shell")
    shell.AppActivate("Command Prompt")

    while True:
        shell.SendKeys("{NUMLOCK}")
        shell.SendKeys("{NUMLOCK}")
        time.sleep(1)


if __name__ == "__main__":
    window = tkinter.Tk()
    # Code to add widgets will go here...
    window.title("Always On V1.0")
    # window.geometry('600x200')

    window.config(bg='white')

    # lbl = tkinter.Label(window, text="Test",width=30,bg='white')
    # lbl.config(font=('Arial 18 bold'))
    # lbl.grid(column=1, row=0)
    lblLoc = tkinter.Label(window, text="", bg='white')
    btn = tkinter.Button(window, text="Always OF", command=always_on_pc, width=20, height=2, bg='red', fg='white')

    btn.grid(column=1, row=1)
    lblLoc.grid(column=1, row=4)
    lblTeam = tkinter.Label(window, text="ioTech.ro", bg='white')
    lblTeam.grid(column=1, row=6)
    window.mainloop()
