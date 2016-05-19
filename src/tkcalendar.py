#!/usr/bin/env python3
#Author : Camilo Olarte|colarte@telesat.com.co|Sept.2003
#Modifier: Felix Lu | lugh82@gmail.com | Oct.2011

import calendar
import tkinter as tk
import time

calendar.setfirstweekday(calendar.SUNDAY)
year = time.localtime()[0]
month = time.localtime()[1]
day = time.localtime()[2]
DATE_DELIMITER = '-'
strdate = (str(year) + DATE_DELIMITER + \
    '{0:0>2}'.format(str(month)) + DATE_DELIMITER + \
    '{0:0>2}'.format(str(day)))

fntTitle = ("Times", 10, 'bold')
fntHeader = ("Times", 11)
fntCal = ("Times", 10)

lang = "zh"
# lang = "zh"
# else lang = "en"

if lang == "zh":
    # Chinese Options
    strtitle = "选择日期"
    strdays= "日  一  二  三  四  五  六"
    dictmonths = {'1': '一月', '2': '二月', '3': '三月', '4': '四月', '5': '五月',
        '6': '六月', '7': '七月', '8': '八月', '9': '九月', '10': '十月',
        '11': '十一月', '12': '十二月'}
else:
    # English Options
    strtitle = "Calendar"
    strdays = "Su  Mo  Tu  We  Th  Fr  Sa"
    dictmonths = {'1': 'Jan', '2': 'Feb', '3': 'Mar', '4': 'Apr', '5': 'May',
        '6': 'Jun', '7': 'Jul', '8': 'Aug', '9': 'Sep', '10': 'Oct',
        '11': 'Nov', '12': 'Dec'}

##############################################
#  BEGIN CLASS

class tkCalendar:

    def __init__(self, master, arg_year, arg_month, arg_day,
        arg_parent_updatable_var):
        self.update_var = arg_parent_updatable_var
        top = self.top = tk.Toplevel(master)
        try:
            self.intmonth = int(arg_month)
        except:
            self.intmonth = int(1)
        self.canvas =tk.Canvas(top, width=200, height=220,
            relief=tk.RIDGE, background="white", borderwidth=1)
        self.canvas.create_rectangle(0, 0, 303, 30, fill="#a4cae8", width=0)
        self.canvas.create_text(100, 17, text=strtitle, font=fntTitle,
            fill="#2024d6")
        stryear = str(arg_year)

        self.year_var = tk.StringVar()
        self.year_var.set(stryear)
        self.lblYear = tk.Label(top, textvariable=self.year_var,
            font=fntHeader, background="white")
        self.lblYear.place(x=85, y=30)

        self.month_var = tk.StringVar()
        strnummonth = str(self.intmonth)
        strmonth = dictmonths[strnummonth]
        self.month_var.set(strmonth)

        self.lblYear = tk.Label(top, textvariable=self.month_var,
            font=fntHeader, background="white")
        self.lblYear.place(x=85, y=50)
        #Variable muy usada
        tagBaseButton = "Arrow"
        self.tagBaseNumber = "DayButton"
        #draw year arrows
        x, y = 30, 43
        tagThisButton = "leftyear"
        tagFinalThisButton = tuple((tagBaseButton, tagThisButton))
        self.fnCreateLeftArrow(self.canvas, x, y, tagFinalThisButton)
        x, y = 160, 43
        tagThisButton = "rightyear"
        tagFinalThisButton = tuple((tagBaseButton, tagThisButton))
        self.fnCreateRightArrow(self.canvas, x, y, tagFinalThisButton)
        #draw month arrows
        x, y = 40, 63
        tagThisButton = "leftmonth"
        tagFinalThisButton = tuple((tagBaseButton, tagThisButton))
        self.fnCreateLeftArrow(self.canvas, x, y, tagFinalThisButton)
        x, y = 150, 63
        tagThisButton = "rightmonth"
        tagFinalThisButton = tuple((tagBaseButton, tagThisButton))
        self.fnCreateRightArrow(self.canvas, x, y, tagFinalThisButton)
        #Print days
        self.canvas.create_text(100, 90, text=strdays, font=fntHeader)
        self.canvas.pack(expand=1, fill=tk.BOTH)
        self.canvas.tag_bind("Arrow", "<ButtonRelease-1>", self.fnClick)
        self.canvas.tag_bind("Arrow", "<Enter>", self.fnOnMouseOver)
        self.canvas.tag_bind("Arrow", "<Leave>", self.fnOnMouseOut)
        self.fnFillCalendar()

    def bind(self, event, *args):
        self.canvas.bind(event, args)

    def fnCreateRightArrow(self, canv, x, y, strtagname):
        canv.create_polygon(x, y, [[x+0, y-5], [x+10, y-5], [x+10, y-10],
            [x+20, y+0], [x+10, y+10], [x+10, y+5], [x+0, y+5]],
            tags=strtagname, fill="blue", width=0)

    def fnCreateLeftArrow(self, canv, x, y, strtagname):
        canv.create_polygon(x, y, [[x+10, y-10], [x+10, y-5], [x+20, y-5],
            [x+20, y+5], [x+10, y+5], [x+10, y+10]],
            tags=strtagname, fill="blue", width=0)

    def fnClick(self, event):
        owntags = self.canvas.gettags(tk.CURRENT)
        if "rightyear" in owntags:
            intyear = int(self.year_var.get())
            intyear += 1
            stryear = str(intyear)
            self.year_var.set(stryear)
        if "leftyear" in owntags:
            intyear = int(self.year_var.get())
            intyear -= 1
            stryear = str(intyear)
            self.year_var.set(stryear)
        if "rightmonth" in owntags:
            if self.intmonth < 12:
                self.intmonth += 1
                strnummonth = str(self.intmonth)
                strmonth = dictmonths[strnummonth]
                self.month_var.set(strmonth)
            else:
                self.intmonth = 1
                strnummonth = str(self.intmonth)
                strmonth = dictmonths[strnummonth]
                self.month_var.set(strmonth)
                intyear = int(self.year_var.get())
                intyear += 1
                stryear = str(intyear)
                self.year_var.set(stryear)
        if "leftmonth" in owntags:
            if self.intmonth > 1:
                self.intmonth -= 1
                strnummonth = str(self.intmonth)
                strmonth = dictmonths[strnummonth]
                self.month_var.set(strmonth)
            else:
                self.intmonth = 12
                strnummonth = str(self.intmonth)
                strmonth = dictmonths[strnummonth]
                self.month_var.set(strmonth)
                intyear = int(self.year_var.get())
                intyear -= 1
                stryear = str(intyear)
                self.year_var.set(stryear)
        self.fnFillCalendar()

    def fnFillCalendar(self):
        init_x_pos = 20
        arr_y_pos = [110, 130, 150, 170, 190, 210]
        intposarr = 0
        self.canvas.delete("DayButton")
        self.canvas.update()
        intyear = int(self.year_var.get())
        monthcal = calendar.monthcalendar(intyear, self.intmonth)
        for row in monthcal:
            xpos = init_x_pos
            ypos = arr_y_pos[intposarr]
            for item in row:
                stritem = str(item)
                if stritem == "0":
                    xpos += 27
                else:
                    tagNumber = tuple((self.tagBaseNumber, stritem))
                    self.canvas.create_text(xpos, ypos, text=stritem,
                        font=fntCal, tags=tagNumber)
                    xpos += 27
            intposarr += 1
        self.canvas.tag_bind("DayButton", "<ButtonRelease-1>",
            self.fnClickNumber)
        self.canvas.tag_bind("DayButton", "<Enter>", self.fnOnMouseOver)
        self.canvas.tag_bind("DayButton", "<Leave>", self.fnOnMouseOut)

    def fnClickNumber(self, event):
        owntags = self.canvas.gettags(tk.CURRENT)
        for x in owntags:
            if (x == "current") or (x == "DayButton"):
                pass
            else:
                strdate = (str(self.year_var.get()) + DATE_DELIMITER +
                    '{0:0>2}'.format(str(self.intmonth)) + DATE_DELIMITER +
                    '{0:0>2}'.format(str(x)))
                self.update_var.set(strdate)
                self.top.withdraw()

    def fnOnMouseOver(self, event):
        self.canvas.move(tk.CURRENT, 1, 1)
        self.canvas.update()

    def fnOnMouseOut(self, event):
        self.canvas.move(tk.CURRENT, -1, -1)
        self.canvas.update()

#  END CLASS
##############################################

class clsMainFrame(tk.Frame):

    def __init__(self, master):
        self.parent = master
        tk.Frame.__init__(master)
        self.date_var = tk.StringVar()
        self.date_var.set(strdate)
        label = tk.Label(master, textvariable=self.date_var, bg="white")
        label.pack(side="top")
        testBtn = tk.Button(master, text='getdate', command=self.fnCalendar)
        testBtn.pack(side='left')
        exitBtn = tk.Button(master, text='Exit', command=master.destroy)
        exitBtn.pack(side='right')

    def fnCalendar(self):
        tkCalendar(self.parent, year, month, day, self.date_var)

if __name__ == '__main__':
    root =tk.Tk()
    root.title("Calendar")
    Frm = tk.Frame(root)
    clsMainFrame(Frm)
    Frm.pack()
    root.mainloop()
