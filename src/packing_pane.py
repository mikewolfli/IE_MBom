#coding=utf-8
'''
Created on 2017年1月24日

@author: 10256603
'''
from global_list import *
global login_info

        
#分箱程序
class packing_pane(Frame):
    def __init__(self,master=None):
        Frame.__init__(self, master)
        self.grid()
        
        self.createWidgets()
        
    def createWidgets(self):
        operate_pane = Frame(self)
        operate_pane.grid(row=0, column=0,rowspan=2,columnspan=2, sticky=NSEW)
        
        self.wbs_button = Button(operate_pane, text='输入WBS NO')
        self.wbs_button.grid(row=0, column=0, sticky=NSEW)
        
        self.export_pl_button = Button(operate_pane, text='导出装箱单')
        self.export_pl_button.grid(row=0, column=2,  sticky=NSEW)
        
        self.calc_pl_button = Button(operate_pane, text='并箱计算')
        self.calc_pl_button.grid(row=1, column=0,  sticky=NSEW)
        
        self.export_re_button = Button(operate_pane, text='导出并箱结果')
        self.export_re_button.grid(row=1, column=2, sticky=NSEW)
        
        self.subbox_text = scrolledtext.ScrolledText(self, wrap=tk.WORD,height=5)
        self.subbox_text.grid(row=2, column=0, rowspan=2, columnspan=2,sticky=NSEW)
        
        wbs_head = ['WBS no','项目名称','梯号','梯型','载重','速度','门型号','开门形式','门宽','门高']
        pl_head = ['箱号','物料号','是否包装箱','名称','数量','WBS element','备注']
        wbs_cols = ['col1','col2','col3','col4','col5','col6','col7','col8','col9','col10']
        pl_cols=['col1','col2','col3','col4','col5','col6','col7']
        
        style = ttk.Style()
        style.configure("Treeview", font=('TkDefaultFont', 10))
        style.configure("Treeview.Heading", font=('TkDefaultFont', 9))  
        self.wbs_list = ttk.Treeview(self, show='headings', columns=wbs_cols, selectmode='browse')
        self.wbs_list.grid(row=4, column=0, rowspan=8, columnspan=2, sticky='nsew')
        
        self.wbs_list.heading('#0', text='')
        
        for i in range(0,10):
            self.wbs_list.heading(wbs_cols[i], text=wbs_head[i])
            
        self.wbs_list.column('col1', width=80, anchor='w')
        self.wbs_list.column('col2', width=200, anchor='w')
        self.wbs_list.column('col3', width=60, anchor='w')
        self.wbs_list.column('col4', width=80, anchor='w')
        self.wbs_list.column('col5', width=60, anchor='w')
        self.wbs_list.column('col6', width=60, anchor='w')
        self.wbs_list.column('col7', width=80, anchor='w')
        self.wbs_list.column('col8', width=100, anchor='w')
        self.wbs_list.column('col9', width=80, anchor='w')
        self.wbs_list.column('col10', width=80, anchor='w')   
        
        wbs_ysb = ttk.Scrollbar(self, orient='vertical', command=self.wbs_list.yview)
        wbs_xsb = ttk.Scrollbar(self, orient='horizontal', command=self.wbs_list.xview)        
        wbs_ysb.grid(row=4, column=2, rowspan=8, sticky='ns')
        wbs_xsb.grid(row=12, column=0, columnspan=2, sticky='ew')               
        
        self.wbs_list.configure(yscroll=wbs_ysb.set, xscroll=wbs_xsb.set)
             
        self.pl_tree = ttk.Treeview(self, columns=pl_cols)
        self.pl_tree.grid(row=0, column=3, rowspan=12, columnspan=4, sticky='nsew')
        
        self.pl_tree.heading('#0', text='')
        for i in range(0,7):
            self.pl_tree.heading(pl_cols[i], text=pl_head[i])
            
        self.pl_tree.column('#0', width=40, anchor='w')
        self.pl_tree.column('col1', width=80, anchor='w')
        self.pl_tree.column('col2', width=100, anchor='w')
        self.pl_tree.column('col3', width=100, anchor='w')
        self.pl_tree.column('col4', width=150, anchor='w')
        self.pl_tree.column('col5', width=100, anchor='w')
        self.pl_tree.column('col6', width=100, anchor='w')
        self.pl_tree.column('col7', width=200, anchor='w')        
         
        pl_ysb = ttk.Scrollbar(self, orient='vertical', command=self.pl_tree.yview)
        pl_xsb = ttk.Scrollbar(self, orient='horizontal', command=self.pl_tree.xview)        
        pl_ysb.grid(row=0, column=7, rowspan=12, sticky='ns')
        pl_xsb.grid(row=12, column=3, columnspan=4, sticky='ew')               
        
        self.pl_tree.configure(yscroll=pl_ysb.set, xscroll=pl_xsb.set)
        
        self.rowconfigure(10, weight=1)
        self.columnconfigure(1, weight=1)
        
         