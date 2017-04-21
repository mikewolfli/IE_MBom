# coding=utf-8
'''
Created on 2017年1月24日

@author: 10256603
'''
from global_list import *
global login_info

logger = logging.getLogger()

merge_op = ['0120','0130','0150','0190','0200','0220','0520']

xlsx_header = ['WBS Element','Material','Plant','BOM Usage','BOM item','Item Category',
               'Material','Material Description','WBS Element','Activity number in network and standard',
               'Requirement Date','Requirement QTY','Withdraw Qty','Full Issue','Item text line 2','Item text line 1']



def cell2str(val):
    if (val is None) or (val == 'N') or (val == '无') or (val=='N/A'):
        return ''
    else:
        return str(val).strip()
    
# 分箱程序
class packing_pane(Frame):
    wbses = []
    data_thread = None
    wbs_bom = {}
    im_method = -1
    
    prj_info_st = {}
    prj_para_st = {}
    
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.grid()

        self.createWidgets()

    def createWidgets(self):
        operate_pane = Frame(self)
        operate_pane.grid(row=0, column=0, rowspan=2,
                          columnspan=2, sticky=NSEW)

        self.wbs_button = Button(operate_pane, text='导入WBS BOM')
        self.wbs_button.grid(row=0, column=0, sticky=NSEW)
        self.wbs_button['command'] = self.wbs_bom_import

        self.export_pl_button = Button(operate_pane, text='导出装箱单')
        self.export_pl_button.grid(row=0, column=2,  sticky=NSEW)

        self.calc_pl_button = Button(operate_pane, text='并箱计算')
        self.calc_pl_button.grid(row=1, column=0,  sticky=NSEW)

        self.export_re_button = Button(operate_pane, text='导出并箱结果')
        self.export_re_button.grid(row=1, column=2, sticky=NSEW)
        
        self.is_sap_interface= BooleanVar()

        check_is_interface = Checkbutton(operate_pane, text="BOM由SAP接口导入", variable=self.is_sap_interface,
                                     onvalue=True, offvalue=False)
        check_is_interface.grid(row=0, column=3, sticky=EW)
        
        self.is_self_prud = BooleanVar()
        check_is_self_prod = Checkbutton(operate_pane, text="厅门自制", variable=self.is_self_prud,
                                         onvalue=True, offvalue=False)
        check_is_self_prod.grid(row=1, column=3,  sticky=W)

        self.subbox_text = scrolledtext.ScrolledText(
            self, wrap=tk.WORD, height=5)
        self.subbox_text.grid(row=2, column=0, rowspan=2,
                              columnspan=2, sticky=NSEW)

        wbs_head = ['WBS no', '项目名称', '梯号', '梯型',
                    '载重', '速度', '厅门型号','轿门型号', '开门形式', '门宽', '门高']
        pl_head = ['箱号', '物料号', '是否包装箱', '名称', '数量', 'WBS element', '备注']
        wbs_cols = ['col1', 'col2', 'col3', 'col4', 'col5',
                    'col6', 'col7', 'col8', 'col9', 'col10','col11']
        pl_cols = ['col1', 'col2', 'col3', 'col4', 'col5', 'col6', 'col7']

        style = ttk.Style()
        style.configure("Treeview", font=('TkDefaultFont', 10))
        style.configure("Treeview.Heading", font=('TkDefaultFont', 9))
        self.wbs_list = ttk.Treeview(
            self, show='headings', columns=wbs_cols, selectmode='browse')
        self.wbs_list.grid(row=4, column=0, rowspan=8,
                           columnspan=2, sticky='nsew')

        self.wbs_list.heading('#0', text='')

        for i in range(0, 10):
            self.wbs_list.heading(wbs_cols[i], text=wbs_head[i])

        self.wbs_list.column('col1', width=80, anchor='w')
        self.wbs_list.column('col2', width=200, anchor='w')
        self.wbs_list.column('col3', width=60, anchor='w')
        self.wbs_list.column('col4', width=80, anchor='w')
        self.wbs_list.column('col5', width=60, anchor='w')
        self.wbs_list.column('col6', width=60, anchor='w')
        self.wbs_list.column('col7', width=80, anchor='w')
        self.wbs_list.column('col8', width=80, anchor='w')
        self.wbs_list.column('col9', width=100, anchor='w')
        self.wbs_list.column('col10', width=80, anchor='w')
        self.wbs_list.column('col11', width=80, anchor='w')

        wbs_ysb = ttk.Scrollbar(self, orient='vertical',
                                command=self.wbs_list.yview)
        wbs_xsb = ttk.Scrollbar(self, orient='horizontal',
                                command=self.wbs_list.xview)
        wbs_ysb.grid(row=4, column=2, rowspan=8, sticky='ns')
        wbs_xsb.grid(row=12, column=0, columnspan=2, sticky='ew')

        self.wbs_list.configure(yscroll=wbs_ysb.set, xscroll=wbs_xsb.set)

        self.pl_tree = ttk.Treeview(self, columns=pl_cols)
        self.pl_tree.grid(row=0, column=3, rowspan=12,
                          columnspan=4, sticky='nsew')

        self.pl_tree.heading('#0', text='')
        for i in range(0, 7):
            self.pl_tree.heading(pl_cols[i], text=pl_head[i])

        self.pl_tree.column('#0', width=40, anchor='w')
        self.pl_tree.column('col1', width=80, anchor='w')
        self.pl_tree.column('col2', width=100, anchor='w')
        self.pl_tree.column('col3', width=100, anchor='w')
        self.pl_tree.column('col4', width=150, anchor='w')
        self.pl_tree.column('col5', width=100, anchor='w')
        self.pl_tree.column('col6', width=100, anchor='w')
        self.pl_tree.column('col7', width=200, anchor='w')

        pl_ysb = ttk.Scrollbar(self, orient='vertical',
                               command=self.pl_tree.yview)
        pl_xsb = ttk.Scrollbar(self, orient='horizontal',
                               command=self.pl_tree.xview)
        pl_ysb.grid(row=0, column=7, rowspan=12, sticky='ns')
        pl_xsb.grid(row=12, column=3, columnspan=4, sticky='ew')

        self.pl_tree.configure(yscroll=pl_ysb.set, xscroll=pl_xsb.set)

        log_pane = Frame(self)

        self.log_label = Label(log_pane)
        self.log_label["text"] = "操作记录"
        self.log_label.grid(row=0, column=0, sticky=W)

        self.log_text = scrolledtext.ScrolledText(log_pane, state='disabled',height=8)
        self.log_text.config(font=('TkFixedFont', 10, 'normal'))
        self.log_text.grid(row=1, column=0, columnspan=2, sticky=EW)
        log_pane.rowconfigure(1, weight=1)
        log_pane.columnconfigure(1, weight=1)

        log_pane.grid(row=13, column=0, columnspan=8, sticky=NSEW)

        # Create textLogger
        text_handler = TextHandler(self.log_text)
        # Add the handler to logger

        logger.addHandler(text_handler)
        logger.setLevel(logging.INFO)

        self.rowconfigure(10, weight=1)
        self.columnconfigure(1, weight=1)
        
    def display_wbs_info(self):
        for row in self.wbs_list.get_children():
            self.wbs_list.delete(row)
            
        for w in self.wbses:
            line = []
            line.append(w)
            s_str = self.prj_info_st[w]['POST1']
            line.append(s_str)
            s_str = self.prj_para_st[w]['TC000']
            line.append(s_str)
            s_str = self.prj_para_st[w]['TC001']
            line.append(s_str)
            s_str = self.prj_para_st[w]['TC002']
            line.append(s_str)
            s_str = self.prj_para_st[w]['TC003']
            line.append(s_str)
            s_str = self.prj_para_st[w]['TC041']
            line.append(s_str)
            s_str = self.prj_para_st[w]['TC035']
            line.append(s_str)
            s_str = self.prj_para_st[w]['TC036']
            line.append(s_str)
            s_str = self.prj_para_st[w]['TC037']
            line.append(s_str)
            s_str = self.prj_para_st[w]['TC038']
            line.append(s_str)
            self.wbs_list.insert('', END, values=line)
                 
    def wbs_bom_import(self):
        self.wbses=[]
        self.wbs_bom={}
        if self.is_sap_interface.get():
            self.im_method = 0
            self.get_bom_from_sap()
        else:
            self.im_method = 1
            self.read_bom_from_file()
        
    def get_bom_from_sap(self):
        d = ask_list("WBS NO输入", 3)
        if not d:
            return

        self.wbses = d
        
        if not self.check_wbs_in_same_prj():
            messagebox.showerror('error', '分箱操作只能针对同一项目下进行!')
            return

        if self.data_thread is not None and self.data_thread.is_alive():
            messagebox.showinfo('提示', '表单刷新线程正在后台刷新列表，请等待完成后再点击!')
            return        

        self.data_thread = refresh_thread(self)
        self.data_thread.setDaemon(True)
        self.data_thread.start()        
    
    def read_bom_from_file(self):
        self.file_list = filedialog.askopenfilenames(
            title="BOM文件", filetypes=[('excel file', '.xlsx'), ('excel file', '.xlsm')])
        
        if not self.file_list :
            return
        
        self.wbses = []
        self.wbs_bom={}
        
        if self.data_thread is not None and self.data_thread.is_alive():
            messagebox.showinfo('提示', '表单刷新线程正在后台刷新列表，请等待完成后再点击!')
            return        

        self.data_thread = refresh_thread(self)
        self.data_thread.setDaemon(True)
        self.data_thread.start()  
        
    def get_from_file(self):
        j=0
        for file in self.file_list:
            j = self.read_file(file,j)
            
        if not self.check_wbs_in_same_prj():
            messagebox.showerror('error', '分箱操作只能针对同一项目下进行!')
            return
             
    def read_file(self, file, j):
        logger.info('正在读取文件...')
        wb = load_workbook(file, read_only=True, data_only=True)
        sheetnames = wb.get_sheet_names()
        
        if len(sheetnames)==0:
            return
        
        ws = wb.get_sheet_by_name('Sheet1')
        
        i_len = len(xlsx_header)
        
        i=1
        col_act = xlsx_header.index('Activity number in network and standard')
        wbs=''
        
        logger.info('正在筛选BOM 清单...')
        while 1:
            i+=1
            
            if wbs != cell2str(ws.cell(row=i, column=1).value):
                wbs = cell2str(ws.cell(row=i, column=1).value)
                if len(wbs)!=0 and wbs not in self.wbses:
                    self.wbses.append(wbs)
            
            if len(wbs)==0:
                break
            
            act = cell2str(ws.cell(row=i, column=col_act+1).value)
            
            if act not in merge_op: 
                continue
            
            line = {}
            for k in range(1, i_len+1):
                line[xlsx_header[k-1]]=cell2str(ws.cell(row=i, column=k).value)
                
            j+=1
            self.wbs_bom[j]=line  
            
        return j 
                                
    def check_wbs_in_same_prj(self):
        if len(self.wbses)==0:
            return False
        
        prj = self.wbses[0][:10]
        for wbs in self.wbses:
            if prj != wbs[:10]:
                return False
        
        return True
            
    def refresh(self):
        if  self.im_method == 0:
            self.get_from_sap()
            
        if self.im_method == 1:
            self.get_from_file()
            
        self.get_unit_info_from_sap()
        
        self.display_wbs_info()
                      
    def get_unit_info_from_sap(self):
        self.prj_info_st={}
        self.prj_para_st = {}
        
        logger.info("正在登陆SAP...")
        config = ConfigParser()
        config.read('sapnwrfc.cfg')
        para_conn = config._sections['connection']
        para_conn['user'] = base64.b64decode(para_conn['user']).decode()
        para_conn['passwd'] = base64.b64decode(para_conn['passwd']).decode()


        try:
            logger.info("正在连接SAP...")
            conn = pyrfc.Connection(**para_conn)

            self.get_unit_info(conn)

        except pyrfc.CommunicationError:
            logger.error("无法连接服务器")
            return -1
        except pyrfc.LogonError:
            logger.error("无法登陆，帐户密码错误！")
            return -1
        except (pyrfc.ABAPApplicationError, pyrfc.ABAPRuntimeError):
            logger.error("函数执行错误。")
            return -1

        conn.close()
            
    def get_unit_info(self, conn):
        l_wbs = len(self.wbses)
        if l_wbs==0:
            return
        
        imp = []

        if l_wbs > 0:
            for w in self.wbses:
                line = dict(POSID=w)
                imp.append(line)
        elif l_wbs == 0:
            return

        logger.info("正在调用RFC函数(ZAP_PS_PROJECT_INFO)...")
        result = conn.call('ZAP_PS_PROJECT_INFO',
                               IT_CE_POSID=imp, CE_WERKS=login_info['plant'])

        logger.info("正在分析并获取项目基本信息...")

        for re in result['OT_PROJ']:
            wbs = format_wbs_no(re['POSID'])

            self.prj_info_st[wbs] = re
            
        logger.info("项目基本信息获取完成")

        logger.info("正在分析并获取项目参数信息...")
              
        for re in result['OT_CONF']:               
            if dict_has_key(self.prj_para_st, format_wbs_no(re['POSID'])):
                if not dict_has_key(self.prj_para_st[format_wbs_no(re['POSID'])], re['ATNAM']):
                    self.prj_para_st[format_wbs_no(re['POSID'])][re['ATNAM']] = re['ATWTB']
            else:
                self.prj_para_st[format_wbs_no(re['POSID'])] = {}
                self.prj_para_st[format_wbs_no(re['POSID'])][re['ATNAM']] = re['ATWTB']
                      
        logger.info("项目参数信息获取完成。")   
        
        print(self.prj_info_st)
        print(self.prj_para_st)    
              
    def get_from_sap(self):
        logger.info("正在登陆SAP...")
        config = ConfigParser()
        config.read('sapnwrfc.cfg')
        para_conn = config._sections['connection']
        para_conn['user'] = base64.b64decode(para_conn['user']).decode()
        para_conn['passwd'] = base64.b64decode(para_conn['passwd']).decode()


        try:
            logger.info("正在连接SAP...")
            conn = pyrfc.Connection(**para_conn)

            self.refresh_wbs_bom(conn)

        except pyrfc.CommunicationError:
            logger.error("无法连接服务器")
            return -1
        except pyrfc.LogonError:
            logger.error("无法登陆，帐户密码错误！")
            return -1
        except (pyrfc.ABAPApplicationError, pyrfc.ABAPRuntimeError):
            logger.error("函数执行错误。")
            return -1

        conn.close()

        return 1
    
    def refresh_wbs_bom(self, conn):
        imp = []

        for w in self.wbses:
            line = dict(POSID=w)
            imp.append(line)

        logger.info("正在调用RFC函数(ZAP_PS_WBSBOM_INFO)...")
        result = conn.call('ZAP_PS_WBSBOM_INFO',
                           IT_CE_WBSBOM=imp, CE_WERKS=login_info['plant'])

        i = 0

        for re in result['OT_WBSBOM']:
            if  re['VORNR'] not in merge_op:
                continue
            
            i += 1
            wbs = re['POSID'] 
            
            re['PSPEL']= act_to_wbs_element(wbs, re['VORNR'])      
                     
            self.wbs_bom[i] = re
