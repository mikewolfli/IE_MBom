# coding=utf-8
'''
Created on 2017年1月24日

@author: 10256603
'''
from global_list import *
global login_info

logger = logging.getLogger()

merge_op = ['0120', '0130', '0150', '0190', '0200', '0220', '0520']

xlsx_header = ['WBS Element', 'Material', 'Plant', 'BOM Usage', 'BOM item', 'Item Category',
               'Material', 'Material Description', 'WBS Element', 'Activity number in network and standard',
               'Requirement Date', 'Requirement QTY', 'Withdraw Qty', 'Full Issue', 'Item text line 2', 'Item text line 1']

export_head = ['wbs no','箱号','物料号','数量','wbs element','备注','物料名称']
old_packing_head = ['wbs_no','箱号','物料号']
box_wbs_head = ['wbs_no', '整梯物料','物料号','数量','RP', '物料名称']

def cell2str(val):
    if (val is None) or (val == 'N') or (val == '无') or (val == 'N/A'):
        return ''
    else:
        return str(val).strip()


def id_to_box(id1):
    s = id1.strip()
    if len(s) != 5:
        return ''

    s_s = s[0:2]
    s_e = s[2:]

    a = str(int(s_s))
    b = str(int(s_e))

    return a + '#-' + b

    
def box_to_id(box):
    box = box.strip()
    pos = box.find('#')

    if pos < 0:
        return ''

    s_s = box[0:pos]

    if len(s_s) == 1:
        s_s = '0' + s_s
    elif len(s_s) == 0:
        return ''

    s_e = box[pos + 2:]

    if len(s_e) == 1:
        s_e = '00' + s_e
    elif len(s_e) == 2:
        s_e = '0' + s_e
    elif len(s_e) == 0:
        s_e = '000'

    return s_s + s_e


def boxid_add(b_id, pl):  # 05001 + 2 = 05003
    b_id = b_id.strip()

    if len(b_id) != 5:
        return b_id

    id1 = int(b_id[2:])

    id1 = id1 + pl

    s_s = b_id[0:2]

    s_e = str(id1)

    if len(s_e) == 1:
        s_e = '00' + s_e
    elif len(s_e) == 2:
        s_e = '0' + s_e
    else:
        s_e = '000'

    return s_s + s_e


def get_chinese(st):
    p = re.compile(u'[\u4e00-\u9fa5]')
    i = -1
    for k in st:
        res = re.match(p, k)
        if res is None:
            break
        i += 1

    if i < 0:
        return None
    else:
        return st[:i + 1]

# 分箱程序


class packing_pane(Frame):
    wbses = []
    data_thread = None
    wbs_bom = {}
    packing_bom = {}
    im_method = -1

    prj_info_st = {}
    prj_para_st = {}

    door_type_group = {}

    boxes_mat_info = {}
    merge_boxes = {}
    pre_fill_boxes = {}
    wbs_bom_boxes = {}

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
        self.export_pl_button['command'] = self.export_packing_bom

        self.calc_pl_button = Button(operate_pane, text='并箱计算')
        self.calc_pl_button.grid(row=1, column=0,  sticky=NSEW)
        self.calc_pl_button['command'] = self.calc_pl

        #self.export_re_button = Button(operate_pane, text='导出并箱结果')
        #self.export_re_button.grid(row=1, column=2, sticky=NSEW)

        self.is_sap_interface = BooleanVar()

        check_is_interface = Checkbutton(operate_pane, text="BOM由SAP接口导入", variable=self.is_sap_interface,
                                         onvalue=True, offvalue=False)
        check_is_interface.grid(row=0, column=3, sticky=EW)
        self.is_sap_interface.set(True)

        self.is_self_prud = BooleanVar()
        check_is_self_prod = Checkbutton(operate_pane, text="厅门自制", variable=self.is_self_prud,
                                         onvalue=True, offvalue=False)
        check_is_self_prod.grid(row=1, column=3,  sticky=W)

        self.is_wooden = BooleanVar()
        check_is_wooden = Checkbutton(
            operate_pane, text="5#木箱", variable=self.is_wooden)
        check_is_wooden.grid(row=1, column=4, sticky=W)

        self.subbox_text = scrolledtext.ScrolledText(
            self, wrap=tk.WORD, height=5, width=10)
        self.subbox_text.grid(row=2, column=0, rowspan=2,
                              columnspan=3, sticky=NSEW)

        wbs_head = ['WBS no', '项目名称', '梯号', '梯型',
                    '载重', '速度', 'Stops','厅门型号', '轿门型号', '开门形式', '门宽', '门高']
        pl_head = ['箱号', '物料号', '是否包装箱', '名称', '数量', 'WBS element', '备注']
        wbs_cols = ['col1', 'col2', 'col3', 'col4', 'col5',
                    'col6', 'col7', 'col8', 'col9', 'col10', 'col11','col12']
        pl_cols = ['col1', 'col2', 'col3', 'col4', 'col5', 'col6', 'col7']

        style = ttk.Style()
        style.configure("Treeview", font=('TkDefaultFont', 10))
        style.configure("Treeview.Heading", font=('TkDefaultFont', 9))
        self.wbs_list = ttk.Treeview(
            self, show='headings', columns=wbs_cols, selectmode='browse')
        self.wbs_list.grid(row=4, column=0, rowspan=8,
                           columnspan=2, sticky='nsew')

        self.wbs_list.heading('#0', text='')

        for i in range(0, 12):
            self.wbs_list.heading(wbs_cols[i], text=wbs_head[i])

        self.wbs_list.column('col1', width=150, anchor='w')
        self.wbs_list.column('col2', width=100, anchor='w')
        self.wbs_list.column('col3', width=60, anchor='w')
        self.wbs_list.column('col4', width=80, anchor='w')
        self.wbs_list.column('col5', width=60, anchor='w')
        self.wbs_list.column('col6', width=60, anchor='w')
        self.wbs_list.column('col7', width=80, anchor='w')
        self.wbs_list.column('col8', width=80, anchor='w')
        self.wbs_list.column('col9', width=80, anchor='w')
        self.wbs_list.column('col10', width=100, anchor='w')
        self.wbs_list.column('col11', width=80, anchor='w')
        self.wbs_list.column('col12', width=80, anchor='w')


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
        self.pl_tree.column('col1', width=150, anchor='w')
        self.pl_tree.column('col2', width=100, anchor='w')
        self.pl_tree.column('col3', width=100, anchor='w')
        self.pl_tree.column('col4', width=150, anchor='w')
        self.pl_tree.column('col5', width=100, anchor='e')
        self.pl_tree.column('col6', width=150, anchor='w')
        self.pl_tree.column('col7', width=300, anchor='w')

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

        self.log_text = scrolledtext.ScrolledText(
            log_pane, state='disabled', height=8)
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
            
            s_str = self.prj_para_st[w]['TC005']
            line.append(s_str)
            s_str = self.prj_para_st[w]['TC041']
            line.append(s_str)
        
            s_str = self.prj_para_st[w]['TC035']
            line.append(s_str)
            s_str = self.prj_para_st[w]['TC036']
            line.append(s_str)
            s_str = self.prj_para_st[w]['TC037']
            line.append(s_str)
            self.wbs_list.insert('', END, values=line)

    def wbs_bom_import(self):
        self.wbses = []
        self.wbs_bom = {}
        self.door_type_group = {}
        self.packing_bom={}
        self.merge_boxes={}
        self.subbox_text.delete('1.0',END)
        self.wbs_bom_boxes = {}

        for row in self.wbs_list.get_children():
            self.wbs_list.delete(row)
            
        for row in self.pl_tree.get_children():
            self.pl_tree.delete(row)

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

        if not self.file_list:
            return

        self.wbses = []
        self.wbs_bom = {}

        if self.data_thread is not None and self.data_thread.is_alive():
            messagebox.showinfo('提示', '表单刷新线程正在后台刷新列表，请等待完成后再点击!')
            return

        self.data_thread = refresh_thread(self)
        self.data_thread.setDaemon(True)
        self.data_thread.start()

    def get_from_file(self):
        j = 0
        for file in self.file_list:
            j = self.read_file(file, j)

        if not self.check_wbs_in_same_prj():
            messagebox.showerror('error', '分箱操作只能针对同一项目下进行!')
            return

    def read_file(self, file, j):
        logger.info('正在读取文件...')
        wb = load_workbook(file, read_only=True, data_only=True)
        sheetnames = wb.get_sheet_names()

        if len(sheetnames) == 0:
            return 0

        ws = wb.get_sheet_by_name('Sheet1')

        i = 1
        col_act = xlsx_header.index('Activity number in network and standard')
        wbs = ''

        logger.info('正在筛选BOM 清单...')
        while 1:
            i += 1

            if wbs != cell2str(ws.cell(row=i, column=1).value):
                wbs = cell2str(ws.cell(row=i, column=1).value)
                if len(wbs) != 0 and wbs not in self.wbses:
                    self.wbses.append(wbs)

            if len(wbs) == 0:
                break

            act = cell2str(ws.cell(row=i, column=col_act + 1).value)

            if act not in merge_op:
                continue

            line = {}

            line['wbs_no'] = cell2str(ws.cell(row=i, column=1).value)
            line['mat_no'] = cell2str(ws.cell(row=i, column=7).value).strip()
            line['mat_name'] = cell2str(ws.cell(row=i, column=8).value)
            line['wbs_element'] = cell2str(ws.cell(row=i, column=9).value)
            line['qty'] = float(cell2str(ws.cell(row=i, column=12).value))
            line['remarks'] = cell2str(ws.cell(row=i, column=16).value).strip()
            line['box_id'] = ''
            line['is_box_mat'] = False
            line['activity'] = cell2str(ws.cell(row=i, column=10).value)

            j += 1
            self.wbs_bom[j] = line

        return j

    def dict_to_list(self, key, d_bom):
        d_l = [key, ]
        d_l.append(d_bom['mat_no'])
        if d_bom['is_box_mat']:
            s_r = 'Y'
        else:
            s_r = ''
        d_l.append(s_r)
        d_l.append(d_bom['mat_name'])
        d_l.append(d_bom['qty'])
        d_l.append(d_bom['wbs_element'])
        d_l.append(d_bom['remarks'])

        return d_l

    def fill_in_packing_list(self, wbs=None):
        for row in self.pl_tree.get_children():
            self.pl_tree.delete(row)

        if wbs is None:
            wbses = sorted(list(self.packing_bom))

            for w in wbses:
                h = [w, '', '', '', '', '', '']
                w_branch = self.pl_tree.insert('', END, values=h)
                p_items = sorted(self.packing_bom[w].keys())
                for p in p_items:
                    a = [id_to_box(p), '', '', '', '', '', '']
                    p_branch = self.pl_tree.insert(w_branch, END, values=a)
                    items = sorted(self.packing_bom[w][p].keys())
                    for itm in items:
                        item = self.packing_bom[w][p][itm]
                        it = self.dict_to_list(itm, item)
                        self.pl_tree.insert(p_branch, END, values=it)

        else:
            if len(self.packing_bom[wbs]) == 0:
                return

            keys = sorted(self.packing_bom[wbs].keys())

            for key in keys:
                a = [id_to_box(key), '', '', '', '', '', '']
                p_branch = self.pl_tree.insert('', END, values=a)
                items = sorted(self.packing_bom[wbs][key].keys())
                for itm in items:
                    item = self.packing_bom[wbs][key][itm]
                    it = self.dict_to_list(itm, item)
                    self.pl_tree.insert(p_branch, END, values=it)

    def export_packing_bom(self):
        if len(self.packing_bom) == 0:
            logger.error('分箱结构为空!')
            return

        file_str = filedialog.asksaveasfilename(
            title="导出文件", filetypes=[('excel file', '.xlsx')])
        if not file_str:
            return

        if not file_str.endswith(".xlsx"):
            file_str += ".xlsx"
            
        wb=Workbook()
        ws0=wb.worksheets[0]

        ws0.title='装箱清单导入表'
        col_size = len(export_head)
        
        for i in range(col_size):
            ws0.cell(row=1,column=i+1).value=export_head[i]

        keys = sorted(self.packing_bom.keys()) 
        i_pos = 1 
        for key in keys:
            boxes = sorted(self.packing_bom[key])
            for box in boxes:
                items = sorted(self.packing_bom[key][box])
                for item in items:
                    ws0.cell(row=i_pos+1, column=1).value = key
                    ws0.cell(row=i_pos+1, column=2).value = id_to_box(box)
                    ws0.cell(row=i_pos+1,column=3).value = self.packing_bom[key][box][item]['mat_no']
                    ws0.cell(row=i_pos+1, column=4).value = self.packing_bom[key][box][item]['qty']
                    ws0.cell(row=i_pos+1, column=5).value = self.packing_bom[key][box][item]['wbs_element']
                    ws0.cell(row=i_pos+1, column=6).value = self.packing_bom[key][box][item]['remarks']
                    ws0.cell(row=i_pos+1, column=7).value = self.packing_bom[key][box][item]['mat_name']
            
                    i_pos+=1
        
        ws1 = wb.create_sheet()
        ws1.title = '旧装箱单删除表'
        
        col_size = len(old_packing_head)
        for i in range(col_size):
            ws1.cell(row=1, column=i+1).value = old_packing_head[i]
        
        keys = sorted(self.wbs_bom.keys())
        i_pos = 1
        for key in keys:
            ws1.cell(row=i_pos+1, column=1).value = self.wbs_bom[key]['wbs_no']
            ws1.cell(row=i_pos+1, column=3).value = self.wbs_bom[key]['mat_no']
            #ws1.cell(row=i_pos+1,  column=4).value =self.wbs_bom[key]['qty']
            i_pos+=1
            
        ws2 = wb.create_sheet()
        ws2.title = 'WBS BOM增加box list'
        
        col_size = len(box_wbs_head)
        for i in range(col_size):
            ws2.cell(row=1, column=i+1).value = box_wbs_head[i]
        
        keys  = sorted(self.wbs_bom_boxes.keys())
        i_pos =1
        for key in keys:
            k_boxes = sorted(self.wbs_bom_boxes[key].keys())
            for k_box in k_boxes:
                ws2.cell(row=i_pos+1, column=1).value = self.wbs_bom_boxes[key][k_box]['wbs_no']
                ws2.cell(row=i_pos+1, column=2).value = self.wbs_bom_boxes[key][k_box]['p_mat']
                ws2.cell(row=i_pos+1, column=3).value = self.wbs_bom_boxes[key][k_box]['mat_no']
                ws2.cell(row=i_pos+1, column=4).value = self.wbs_bom_boxes[key][k_box]['qty']
                ws2.cell(row=i_pos+1, column=5).value = self.wbs_bom_boxes[key][k_box]['rp']
                ws2.cell(row=i_pos+1, column=6).value = self.wbs_bom_boxes[key][k_box]['mat_name']
                i_pos+=1
                           
        if excel_xlsx.save_workbook(workbook=wb, filename=file_str):
            messagebox.showinfo("输出","成功输出!")        
               
    def add_merge_comments(self):
        if len(self.merge_boxes) == 0:
            return

        self.clear_merge_boxes()
        
        keys = sorted(list(self.merge_boxes))
        
        merg_str = ''
        
        
        for key in keys:
            i_len = len(self.merge_boxes[key])
            s_str = ''
            for i in range(i_len):
                li = self.merge_boxes[key][i]
                if len(s_str)==0:
                    s_str = self.get_merge_comment(li, key)
                else:
                    s_str = s_str +';'+ self.get_merge_comment(li, key)
                
            if len(merg_str)>0:
                merg_str += '\n'+id_to_box(key)+':'+s_str 
            else:
                merg_str = id_to_box(key)+':'+s_str
        
        self.subbox_text.delete('1.0',END)
        self.subbox_text.insert(END, merg_str)   
            
    def get_merge_comment(self, li_dic, b_id):
        li = sorted(list(li_dic))
        i_len = len(li)
        s_str = ''
        s_in = ''
        s_in_wbs=''
        s_qty=''
        s_re = ''
        
        for i in range(0, i_len):
            s_wbs = li[i]
            s_qty =str(int( li_dic[s_wbs]))
            if i ==0:
                s_in = self.get_unit_no(s_wbs)
                s_str = s_in
                s_re = s_qty+'/'+s_in
                s_in_wbs=s_wbs
            else:
                s_str= s_str+','+ self.get_unit_no(s_wbs)
                s_re= s_re+','+s_qty+'/'+self.get_unit_no(s_wbs)
                items = sorted(self.packing_bom[s_wbs][b_id].keys())
                
                for itm in items:
                    self.packing_bom[s_wbs][b_id][itm]['remarks'] += '装于'+s_in              
            
            if i==i_len-1:
                s_str = s_str+'并箱'
                s_re= s_re+'并箱'
                
        if len(s_in_wbs)>0:
            itms = sorted(self.packing_bom[s_in_wbs][b_id].keys())
            for it in itms:
                if len(self.packing_bom[s_in_wbs][b_id][it]['remarks'] )>0:
                    self.packing_bom[s_in_wbs][b_id][it]['remarks'] += (','+s_str)
                else:
                    self.packing_bom[s_in_wbs][b_id][it]['remarks'] += s_str
        
        return s_re
            
    def clear_merge_boxes(self):
        keys = list(self.merge_boxes)
        for key in keys:
            self.merge_boxes[key] = [x for x in self.merge_boxes[key] if len(x)>1 ]
                                
            if len(self.merge_boxes[key])==0:
                self.merge_boxes.pop(key)


    def get_unit_no(self, wbs):
        return self.prj_para_st[wbs]['TC000']

    def check_wbs_in_same_prj(self):
        if len(self.wbses) == 0:
            return False

        prj = self.wbses[0][:10]
        for wbs in self.wbses:
            if prj != wbs[:10]:
                return False

        return True

    def refresh(self):
        if self.im_method == 0:
            self.get_from_sap()

        if self.im_method == 1:
            self.get_from_file()

        if self.im_method == 0 or self.im_method == 1:
            self.get_unit_info_from_sap()
            self.display_wbs_info()

        if self.im_method == 2:
            self.group_door_type()
            self.scan_bom_fill_boxid()
            sum_pl_bom = self.sum_mat_qty()
            self.get_box_mat_qty(sum_pl_bom)
            self.get_doors_qty()
            r_packing_bom = self.reverse_packing_bom()
            self.split_boxes_bom(r_packing_bom)
            self.add_merge_comments()
            self.fill_in_packing_list()

    def get_doors_qty(self):
        self.doors_qty = {}
        keys = sorted(self.packing_bom.keys())

        for key in keys:
            qty = 0.0
            boxes = sorted(self.packing_bom[key].keys())
            if '05002' not in boxes:
                return

            lines = sorted(self.packing_bom[key]['05002'].keys())

            for li in lines:
                if '悬挂组件' in self.packing_bom[key]['05002'][li]['mat_name']:
                    qty += self.packing_bom[key]['05002'][li]['qty']

            self.doors_qty[key] = qty

    def reverse_packing_bom(self):
        keys = sorted(self.door_type_group.keys())
        r_packing_bom = {}
        for key in keys:
            r_packing_bom[key] = {}
            wbses = sorted(self.door_type_group[key][1])
            boxes = sorted(self.boxes_mat_info[key].keys())
            for box in boxes:
                r_packing_bom[key][box] = {}
                for wbs in wbses:
                    r_packing_bom[key][box][wbs] = copy.deepcopy(
                        self.packing_bom[wbs][box])

        return r_packing_bom

    def split_boxes_bom(self, r_packing_bom):
        self.merge_boxes = {}
        self.packing_bom = {}

        keys = sorted(self.door_type_group.keys())
        
        for key in keys:
            boxes = sorted(self.boxes_mat_info[key].keys())
            w_list = self.door_type_group[key][1]
            for w in w_list:
                self.packing_bom[w] = {}

            for box in boxes:
                box_content = copy.deepcopy(r_packing_bom[key][box] )
                b_infos = self.boxes_mat_info[key][box]
                self.scan_box_content(box, box_content, b_infos)
        
        #print(self.merge_boxes)
                  
    def add_merge_box(self, line, key, qty):
        if key not in list(line):
            line[key] = qty
        else:
            line[key]+=qty
            
    def scan_box_content(self, b_id, box_content, b_infos):
        #keys = sorted(box_content.keys())
        keys = sorted(list(box_content))
        #len_keys = len(keys)
        
        c_lines = {}
        for key in keys:
            c_line = self.catalog_content(
                box_content[key])  # c_line , catalog_line
            c_lines[key] = c_line

        new_bid = b_id
    
        b_landing_set = False

        if b_id == '05001':
            b_hv = self.check_if_hv(c_lines)

            for key in keys:
                #self.packing_bom[key][new_bid] = {}
                it_keys = sorted(list(box_content[key]))
                for it_key in it_keys:
                    if '厅门装置' in box_content[key][it_key]['mat_name']:
                        if new_bid not in list(self.packing_bom[key]):
                            self.packing_bom[key][new_bid] = {}

                        p_id = len(self.packing_bom[key][new_bid])
                        self.packing_bom[key][new_bid][p_id +
                                                       1] = box_content[key][it_key].copy()

                        box_content[key].pop(it_key)
                        b_landing_set = True

            if b_landing_set:
                for key in keys:
                    if '厅门装置' in c_lines[key].keys():
                        c_lines[key].pop('厅门装置')

                new_bid = boxid_add(new_bid, 3)

        else:
            b_hv = False
        
        if b_hv:
            b_keys = sorted(b_infos.keys(), reverse=True) 
        else:
            b_keys = sorted(b_infos.keys()) 
            
        #pre_fill_boxes={}
        s_item=''
        line_merge_box={}

        for b_key in b_keys:
            b_mat = b_infos[b_key][0]
            b_max = b_infos[b_key][2]
            b_mat_max = b_max
            b_c_max = b_max
            #b_min = b_infos[b_key][1]
            b_qty = b_infos[b_key][3]

            while b_qty > 0:
                box_keys = sorted(list(self.merge_boxes))
                #line_merge_box = []

                for key in keys:
                    # if new_bid not in list(self.packing_bom[key]):
                    #    self.packing_bom[key][new_bid] = {}    
                    if b_c_max !=b_max and b_c_max !=0:
                        b_max = b_c_max
                        
                    if b_hv:
                        items = list(c_lines[key])
                        i_len = len(items)
                        
                        for item in items:
                            if len(s_item)==0:
                                s_item  = item
                            if b_c_max ==0  or  items.index(item)!=0:
                                b_c_max = b_max
                                
                            lies = sorted(list(c_lines[key][item]))
                            rt = int(self.sum_catalog_qty(
                                c_lines[key][item]) / self.doors_qty[key])
                            if rt <= 1:
                                rt = 1
                            elif rt > 1 and rt < 2:
                                rt = 1

                            if len(lies) > 1:
                                li = self.get_min_item(c_lines[key][item])
                            elif len(lies) == 1:
                                li = sorted(list(c_lines[key][item]))[0]
                            else:
                                continue

                            c_qty = int(c_lines[key][item][li] / rt)
                            if b_c_max > c_qty:

                                if new_bid not in list(self.packing_bom[key]):
                                    self.packing_bom[key][new_bid] = {}

                                p_id = len(self.packing_bom[key][new_bid])

                                if p_id == 0 and b_c_max == b_mat_max:
                                    box_line = self.get_box_mat_info(key, b_mat)
                                    self.packing_bom[key][new_bid][p_id +1] = box_line
                                    self.get_box_mat_list(box_line)
                                
                                    p_id += 1

                                if i_len - 1 == items.index(item):
                                    b_c_max = b_c_max - c_qty
                                    
                                self.packing_bom[key][new_bid][p_id +
                                                               1] = box_content[key][li].copy()
                                box_content[key].pop(li)
                                c_lines[key][item].pop(li)

                                #if key not in line_merge_box:
                                #   line_merge_box.append(key)
                                if s_item == item:
                                    self.add_merge_box(line_merge_box, key, c_qty)
                                '''
                                if i_len - 1 == items.index(item):
                                    if new_bid not in list(pre_fill_boxes):
                                        pre_fill_boxes[new_bid] = []
                                        
                                    if key not in pre_fill_boxes[new_bid]:
                                        pre_fill_boxes[new_bid].append(key)
                                '''        
                                if keys.index(key) == len(keys) -1:
                                    #if key not in line_merge_box:
                                    #    line_merge_box.append(key)

                                    if i_len - 1 == items.index(item):
                                        b_max = b_c_max

                                        if new_bid not in box_keys:
                                            self.merge_boxes[new_bid] = []
                                        
                                        if line_merge_box not in self.merge_boxes[new_bid] and len(line_merge_box)>1:
                                            self.merge_boxes[new_bid].append(line_merge_box.copy()) 
                                            
                                        '''
                                        if new_bid in list(pre_fill_boxes):
                                            if key not in pre_fill_boxes[new_bid]:
                                                pre_fill_boxes[new_bid].append(key)
                                            
                                            if pre_fill_boxes[new_bid] not in self.merge_boxes[new_bid]:
                                                self.merge_boxes[new_bid].append(pre_fill_boxes[new_bid].copy())
                                            pre_fill_boxes = {}
                                        else:
                                            if line_merge_box not in self.merge_boxes[new_bid]:
                                                self.merge_boxes[new_bid].append(line_merge_box)  
                                        '''                               

                            else:
                                c_lines[key][item][li] = c_lines[key][item][li] - rt * b_c_max
                                
                                if c_lines[key][item][li] == 0.0:
                                    c_lines[key][item].pop(li)
                                    
                                if new_bid not in list(self.packing_bom[key]):
                                    self.packing_bom[key][new_bid] = {}

                                p_id = len(self.packing_bom[key][new_bid])

                                if p_id == 0 and b_c_max == b_mat_max:
                                    box_line = self.get_box_mat_info(
                                        key, b_mat)
                                    self.packing_bom[key][new_bid][p_id +
                                                                   1] = box_line
                                    self.get_box_mat_list(box_line)
                                    p_id += 1

                                self.packing_bom[key][new_bid][p_id +
                                                               1] = box_content[key][li].copy()
                                self.packing_bom[key][new_bid][p_id +
                                                               1]['qty'] = rt * b_c_max
                                box_content[key][li]['qty'] = box_content[key][li]['qty'] - rt * b_c_max
                                
                                
                                if box_content[key][li]['qty']==0.0:
                                    box_content[key].pop(li)
                                
                                #if key not in line_merge_box:
                                #    line_merge_box.append(key)
                                if s_item == item:
                                    self.add_merge_box(line_merge_box, key, b_c_max)
                                    
                                b_c_max = 0

                                if i_len - 1 == items.index(item):
                                    b_max = b_c_max

                                    if new_bid not in box_keys:
                                        self.merge_boxes[new_bid] = []
                                        
                                    if line_merge_box not in self.merge_boxes[new_bid] and len(line_merge_box)>1:
                                        self.merge_boxes[new_bid].append(line_merge_box.copy()) 
                                        
                                    '''
                                    if new_bid in list(pre_fill_boxes):
                                        if key not in pre_fill_boxes[new_bid]:
                                            pre_fill_boxes[new_bid].append(key)
                                            
                                        if pre_fill_boxes[new_bid] not in self.merge_boxes[new_bid]:
                                            self.merge_boxes[new_bid].append(pre_fill_boxes[new_bid].copy())
                                        pre_fill_boxes = {}
                                    else:
                                        if line_merge_box not in self.merge_boxes[new_bid]:
                                            self.merge_boxes[new_bid].append(line_merge_box)
                                    '''
                    else:
                        items = list(c_lines[key])
                        i_len = len(items)
                        
                        
                        for item in items:
                            if len(s_item)==0:
                                s_item=item
                                
                            if b_c_max ==0  or  items.index(item)!=0:
                                b_c_max = b_max
                            #lies = sorted(list(c_lines[key][item]))
                            rt = int(self.sum_catalog_qty(
                                c_lines[key][item]) / self.doors_qty[key])
                            if rt <= 1:
                                rt = 1
                            elif rt > 1 and rt < 2:
                                rt = 1

                            c_qty = int(self.sum_catalog_qty(
                                c_lines[key][item]) / rt)
                                                        
                            if b_c_max > c_qty:

                                if new_bid not in list(self.packing_bom[key]):
                                    self.packing_bom[key][new_bid] = {}

                                p_id = len(self.packing_bom[key][new_bid])

                                if p_id == 0 and b_c_max == b_mat_max:
                                    box_line = self.get_box_mat_info(
                                        key, b_mat)
                                    self.packing_bom[key][new_bid][p_id +
                                                                   1] = box_line
                                    self.get_box_mat_list(box_line)
                                    p_id += 1
                                #b_c_max = b_c_max - rt*c_qty
                                    
                                c_keys = list(c_lines[key][item])

                                for c_key in c_keys:
                                    self.packing_bom[key][new_bid][p_id +
                                                                   1] = box_content[key][c_key].copy()
                                    b_c_max = b_c_max - int(c_lines[key][item][c_key] / rt)
                                    c_lines[key][item].pop(c_key)
                                    box_content[key].pop(c_key)
                                    p_id += 1

                                c_lines[key].pop(item)

                                #if key not in line_merge_box:
                                #    line_merge_box.append(key)
                                if s_item == item:
                                    self.add_merge_box(line_merge_box, key, c_qty)
                                
                                '''
                                if i_len - 1 == items.index(item):
                                    if new_bid not in list(pre_fill_boxes):
                                        pre_fill_boxes[new_bid] = []
                                    
                                    if key not in pre_fill_boxes[new_bid]:
                                        pre_fill_boxes[new_bid].append(key)
                                '''        
                                if keys.index(key) == len(keys) -1:
                                    if key not in line_merge_box:
                                        line_merge_box.append(key)

                                    if i_len - 1 == items.index(item):
                                        b_max = b_c_max

                                        if new_bid not in box_keys:
                                            self.merge_boxes[new_bid] = []
                                            
                                        if line_merge_box not in self.merge_boxes[new_bid] and len(line_merge_box)>1:
                                            self.merge_boxes[new_bid].append(line_merge_box.copy()) 
                                            
                                        '''
                                        if new_bid in list(pre_fill_boxes):
                                            if key not in pre_fill_boxes[new_bid]:
                                                pre_fill_boxes[new_bid].append(key)
                                            
                                            if pre_fill_boxes[new_bid] not in self.merge_boxes[new_bid]:
                                                self.merge_boxes[new_bid].append(pre_fill_boxes[new_bid].copy())
                                            pre_fill_boxes = {}
                                        else:
                                            if line_merge_box not in self.merge_boxes[new_bid]:
                                                self.merge_boxes[new_bid].append(line_merge_box) 
                                        ''' 

                            else:
                                c_keys = sorted(list(c_lines[key][item]))
                                if new_bid not in list(self.packing_bom[key]):
                                    self.packing_bom[key][new_bid] = {}

                                p_id = len(self.packing_bom[key][new_bid])

                                if p_id == 0 and b_c_max == b_mat_max:
                                    box_line = self.get_box_mat_info(
                                        key, b_mat)
                                    self.packing_bom[key][new_bid][p_id +
                                                                   1] = box_line
                                    self.get_box_mat_list(box_line)
                                    p_id += 1

                                if s_item == item:
                                    self.add_merge_box(line_merge_box, key, b_c_max)
                                  
                                for c_key in c_keys:
                                    if c_lines[key][item][c_key] >= rt * b_c_max:
                                        c_lines[key][item][c_key] = c_lines[key][item][c_key] - rt * b_c_max
                                        
                                        if c_lines[key][item][c_key] == 0.0:
                                            c_lines[key][item].pop(c_key)

                                        self.packing_bom[key][new_bid][p_id +
                                                                       1] = box_content[key][c_key].copy()
                                        self.packing_bom[key][new_bid][p_id +
                                                                       1]['qty'] = rt * b_c_max
                                        box_content[key][c_key]['qty'] = box_content[key][c_key]['qty'] - rt * b_c_max
                                        
                                        if box_content[key][c_key]['qty'] == 0.0:
                                            box_content[key].pop(c_key)
                                            
                                        b_c_max = 0
                                            
                                        break
                                    else:
                                        self.packing_bom[key][new_bid][p_id + 1] = box_content[key][c_key].copy()
                                        b_c_max = b_c_max - int(c_lines[key][item][c_key] / rt)
                                        c_lines[key][item].pop(c_key)
                                        box_content[key].pop(c_key)

                                    p_id += 1

                                #if key not in line_merge_box:
                                #    line_merge_box.append(key)
                                    
                                if i_len - 1 == items.index(item):
                                    b_max = b_c_max
                                    if new_bid not in box_keys:
                                        self.merge_boxes[new_bid] = []
                                        
                                    if line_merge_box not in self.merge_boxes[new_bid] and len(line_merge_box)>1:
                                        self.merge_boxes[new_bid].append(line_merge_box.copy())
                                        
                                    '''
                                    if new_bid in list(pre_fill_boxes):
                                        if key not in pre_fill_boxes[new_bid]:
                                            pre_fill_boxes[new_bid].append(key)
                                            
                                        if pre_fill_boxes[new_bid] not in self.merge_boxes[new_bid]:
                                            self.merge_boxes[new_bid].append(pre_fill_boxes[new_bid].copy())
                                        pre_fill_boxes = {}
                                    else:
                                        if line_merge_box not in self.merge_boxes[new_bid]:
                                            self.merge_boxes[new_bid].append(line_merge_box)
                                    '''

                    if len(c_lines[key]) == 0:
                        c_lines.pop(key)

                    if len(box_content[key]) == 0:
                        box_content.pop(key)
                        keys.remove(key)

                    if b_max == 0:
                        break
                
                if b_max == 0:
                    b_qty -= 1
                    b_max = b_infos[b_key][2]
                    if len(keys)>0:
                        new_bid = b_id
                        line_merge_box={}
                        while new_bid in self.packing_bom[keys[0]].keys():
                            new_bid = boxid_add(new_bid, 3)

                if len(box_content) == 0:
                    break
    
    def get_box_mat_list(self, b_line):
        wbs_no = b_line['wbs_no']
        prod_mat = self.prj_info_st[wbs_no]['MATNR'][-9:]
        b_mat = b_line['mat_no']
        b_wbs = {}
        
        b_wbs['wbs_no'] = wbs_no
        
        b_wbs['mat_no'] = b_line['mat_no']
        b_wbs['p_mat'] = prod_mat
        b_wbs['qty'] = b_line['qty']
        
        if len(b_line['activity'])>0:
            b_wbs['rp'] = 'A' + b_line['activity'][1:]
        else:
            b_wbs['rp'] = ''
            
        b_wbs['mat_name'] = b_line['mat_name']
        
        if wbs_no not in list(self.wbs_bom_boxes):
            self.wbs_bom_boxes[wbs_no] = {}
            
        if b_mat not in list(self.wbs_bom_boxes[wbs_no]):
            self.wbs_bom_boxes[wbs_no][b_mat] = b_wbs
        else:
            self.wbs_bom_boxes[wbs_no][b_mat]['qty'] += b_wbs['qty']
              
    def get_min_item(self, lines):
        keys = sorted(lines.keys())
        i = 0
        for key in keys:
            if i == 0:
                n = key
                f = lines[key]
            elif f > lines[key]:
                n = key
                f = lines[key]

            i += 1

        return n

    def get_rate(self, wbs, k, c_lines, d_qty):
        contents = c_lines[wbs]
        d_q = d_qty[wbs].copy()
        if '厅门装置' in contents.keys():
            q = self.sum_catalog_qty(contents['厅门装置'])
            d_q = d_q - q

        rate = int(self.sum_catalog_qty(contents[k]) / d_q)

        return rate

    def check_if_hv(self, c_lines):
        for key in list(c_lines):
            for li in list(c_lines[key]):
                if len(c_lines[key][li]) > 1:
                    return True

    def sum_catalog_qty(self, items):
        qty = 0.0
        for key in items.keys():
            qty += items[key]

        return qty

    '''        
    def split_boxes_bom(self, q_dic, r_packing_bom):
        self.merge_boxes={}
        keys = sorted(self.door_type_group.keys())
        rate=1
        
        for key in keys:
            
            if q_dic[key][1] == 2* q_dic[key][2]:                          
                rate = 2
            else:
                rate = 1
        
            boxes = sorted(self.boxes_mat_info[key].keys())
            
            self.merge_boxes[key] = {}
            for box in boxes:  
                box_content = copy.deepcopy(r_packing_bom[key][box])
                b_infos = self.boxes_mat_info[key][box]
                                   
                self.scan_box_content(box, box_content, b_infos, rate, self.merge_boxes[key])
    
    def scan_box_content(self, b_id, box_content, b_infos, rate, merge_box):#仅扫描5#-1，
        keys = sorted(box_content.keys())
        box_count = len(b_infos)
        
        b_cont = []      
        q_l = self.min_content(box_content[keys[0]])
        
        min_q = q_l[0]
        b_same = q_l[1]
        
        
        if b_same:
            i=0
            for b in range(1,box_count+1):
                b_max = b_infos[b][2]
                if b_infos[b][3]==0:
                    b_infos.pop(b)
                    
                if len(b_infos)==0:
                    break               
                
                new_bid = b_id
                b_new_bid = True
                
                i_k = 0
                for key in keys:
                    
                    if b_new_bid:
                        me_box = []
                    w_i = 1
                    lines = {}
                    if new_bid not in merge_box.keys():
                        merge_box[new_bid]=[]
                    
                    if b_id=='05001':
                        for it in list(box_content[key]):
                            m_no = box_content[key][it]['mat_no']
                            m_remarks= box_content[key][it]['remarks']
                            
                            if i_k==0 and box_content[key][it]['qty']==min_q:
                                b_cont.append(m_no+m_remarks)
                                                          
                            if m_no+m_remarks in b_cont:
                                l_q = box_content[key][it]['qty'] 
                                if l_q/rate >= b_max:
                                    lines[w_i] = box_content[key][it].copy()
                                    lines[w_i]['qty'] = b_infos[b][2]*rate
                                    box_content[key][it]['qty'] -= b_infos[b][2]*rate
                
                                    if w_i ==1:
                                        i+=1
                                        b_new_bid = True
                                        if b_max < b_infos[b][2]:
                                            if len(me_box)!=0:
                                                me_box.append(key)
                                                merge_box[new_bid].append(me_box)
                                        lines[0]= self.get_box_mat_info(key, b_infos[b][0])
                                        b_max = b_infos[b][2]
                                        b_infos[b][3]-=1
                                else:
                                    lines[w_i] = box_content[key][it].copy()
                                    box_content[key][it]['qty'] = 0.0
                                    if w_i==1:
                                        b_new_bid = False
                                        b_max = b_infos[b][2] - lines[w_i]['qty']
                                        me_box.append(key)
                                    
                                w_i+=1                       
                    else:
                        for it in list(box_content[key]):
                            m_no = box_content[key][it]['mat_no']
                            m_remarks= box_content[key][it]['remarks']
                            
                            if i_k ==0 and box_content[key][it]['qty']==min_q:
                                b_cont.append(m_no+m_remarks)
                                                          
                            if m_no+m_remarks in b_cont:
                                l_q = box_content[key][it]['qty'] 
                                if l_q >= b_max:
                                    lines[w_i] = box_content[key][it].copy()
                                    lines[w_i]['qty'] = b_infos[b][2]
                                    box_content[key][it]['qty'] -= b_infos[b][2]
                
                                    if w_i ==1:
                                        i+=1
                                        b_new_bid = True
                                        if b_max < b_infos[b][2]:
                                            if len(me_box)!=0:
                                                me_box.append(key)
                                                merge_box[new_bid].append(me_box)
                                        lines[0]= self.get_box_mat_info(key, b_infos[b][0])
                                        b_max = b_infos[b][2]
                                        b_infos[b][3]-=1
                                else:
                                    lines[w_i] = box_content[key][it].copy()
                                    box_content[key][it]['qty'] = 0.0
                                    if w_i==1:
                                        b_new_bid = False
                                        b_max = b_infos[b][2] - lines[w_i]['qty']
                                        me_box.append(key)
                                    
                                w_i+=1
                                  
                    self.packing_bom[key][new_bid]=lines
                    
                    if b_new_bid:           
                        new_bid = boxid_add(b_id, (i-1)*3)
                    
                    if b_infos[b][3]==0:
                        b_infos.pop(b)
                        b+=1
                    i_k+=1
        else:
            i=0
            b_new_bid = True
            b_device = False
            new_bid = b_id
            
            i_k=0
            for key in keys:
                if b_new_bid:
                    merge_box[new_bid] = []
                    me_box = []
                w_i = 1
                l_q = 0.0
                lines = {}  
                if new_bid not in merge_box.keys():
                    merge_box[new_bid]=[]
                
                for it in list(box_content[key]):
                    m_no = box_content[key][it]['mat_no']
                    m_name = box_content[key][it]['mat_name']
                    m_remarks= box_content[key][it]['remarks']
                    
                    if i_k==0 and box_content[key][it]['qty']==min_q:
                        b_cont.append(m_no+m_remarks)
                                            
                    if m_no+m_remarks in b_cont:
                        if '厅门装置' in m_name:
                            b_device = True
                            lines[w_i] = box_content[key][it].copy()
                            box_content[key].pop(it)
                        else:
                            if  w_i==1:    
                                l_q += box_content[key][it]['qty']
                                me_box.append(key)
                            
                            if m_no+m_remarks in b_cont:
                                lines[w_i] = box_content[key][it].copy()
                                box_content[key].pop(it)
                    w_i+=1
                
                if b_device:
                    if len(box_content[key])==0:
                        box_content.pop(key)
                        
                self.packing_bom[key][b_id]=lines
                
                
                i_k+=1
                    
                               
            if len(box_content)==0:
                return 
            
            b_merge = False
            if box_count>=1:
                if b_infos[box_count][2]>l_q:
                    box_mat = b_infos[box_count][0]
                    b_max = b_infos[box_count][2]
                    b_infos[box_count][3]-=1
                    if b_infos[box_count][3]==0:
                        b_infos.pop(box_count)
                    box_min_qty=1
                    merge_box[b_id].append(me_box)
                    b_merge = False
                elif b_infos[1][2] >l_q:
                    box_mat = b_infos[1][0]
                    b_max = b_infos[box_count][2]
                    b_infos[1][3]-=1
                    if b_infos[1][3]==0:
                        b_infos.pop(1)
                    box_min_qty=1
                    merge_box[b_id].append(me_box)
                    b_merge = False
                else:
                    box_mat = b_infos[1][0]
                    b_max = b_infos[1][2]
                    box_min_qty = int(l_q/b_infos[1][2])
                    b_infos[1][3]-=box_min_qty
                    b_merge=True
                    
                p_keys = sorted(self.packing_bom.keys())
                i_key = 0;
                
                if b_merge:
                    me_box=[]
                for p_key in p_keys:
                    i_in = self.index_of_pack(b_cont, self.packing_bom[p_key][b_id])
                    i_key += int(self.packing_bom[p_key][b_id][i_in]['qty'])
                    it_no = len(self.packing_bom[p_key][b_id])
                    self.packing_bom[p_key][b_id][it_no] = self.get_box_mat_info(p_key, box_mat)
                    me_box.append(p_key)
                    if i_key==b_max:
                        merge_box[b_id].append(me_box)
                        me_box=[]
                        
            i=0
            for b in range(1, box_count+1):
                b_max = b_infos[b][2]
                if b_infos[b][3]==0:
                    b_infos.pop(b)
                    
                if len(b_infos)==0:
                    break               
                
                new_bid = boxid_add(b_id,3)
                b_new_bid = True
                
                i_k=0
                
                for key in keys:
                    
                    if b_new_bid:
                        merge_box[new_bid] = []
                        me_box = []
                    w_i = 1
                    lines = {}                    

                    for it in list(box_content[key]):
                        m_no = box_content[key][it]['mat_no']
                        m_remarks= box_content[key][it]['remarks']
                          
                        if i_k==0 and box_content[key][it]['qty']==min_q:
                            b_cont.append(m_no+m_remarks)
                                                          
                        if m_no+m_remarks in b_cont:
                            l_q = box_content[key][it]['qty'] 
                            if l_q >= b_max:
                                lines[w_i] = box_content[key][it].copy()
                                lines[w_i]['qty'] = b_infos[b][2]
                                box_content[key][it]['qty'] -= b_infos[b][2]
                
                                if w_i ==1:
                                    i+=1
                                    b_new_bid = True
                                    if b_max < b_infos[b][2]:
                                        if len(me_box)!=0:
                                            me_box.append(key)
                                            merge_box[new_bid].append(me_box)
                                    lines[0]= self.get_box_mat_info(key, b_infos[b][0])
                                    b_max = b_infos[b][2]
                                    b_infos[b][3]-=1
                            else:
                                lines[w_i] = box_content[key][it].copy()
                                box_content[key][it]['qty'] = 0.0
                                if w_i==1:
                                    b_new_bid = False
                                    b_max = b_infos[b][2] - lines[w_i]['qty']
                                    me_box.append(key)
                                    
                            w_i+=1
                                  
                    self.packing_bom[key][new_bid]=lines
                    
                    if b_new_bid:           
                        new_bid = boxid_add(b_id, (i-1)*3)
                    
                    if b_infos[b][3]==0:
                        b_infos.pop(b)
                        b+=1 
                    i_k+=1       
    
    def index_of_pack(self, info, packing):
        for k in list(packing):
            if packing[k]['mat_no']+packing[k]['remarks'] in info:
                return k                        
    def min_content(self, items):
        i=0
        b_same = True
        for l in items.keys():
            if '箱' not in items[l]['mat_name']:
                i+=1
                if i ==1:
                    m_q = items[l]['qty']
                elif m_q > items[l]['qty']:
                    b_same=False
                    m_q = items[l]['qty']
                
        return (m_q, b_same)
    '''

    def catalog_content(self, items):
        res = {}

        for l in items.keys():
            n_index = get_chinese(items[l]['mat_name'])

            if n_index not in list(res):
                res[n_index] = {}

            res[n_index][l] = items[l]['qty']

        return res

    def get_box_mat_info(self, wbs, box):
        box_mat = {}
        try:
            res = mat_info.get(mat_info.mat_no == box)
        except mat_info.DoesNotExist:
            return box_mat

        box_mat['mat_no'] = box
        box_mat['mat_name'] = res.mat_name_cn
        box_mat['qty'] = 1
        box_mat['wbs_no'] = wbs
        box_mat['remarks'] = ''
        box_mat['is_box_mat'] = True
        box_mat['activity'] = str(res.rp).replace('A', '0')
        box_mat['box_id'] = res.box_code_sj
        box_mat['wbs_element'] = act_to_wbs_element(wbs, box_mat['activity'])

        return box_mat

    def get_box_mat_qty(self, s_pl_bom):
        # 1 门板数量， 2 - 悬挂数量 3-特殊 厅门装置数量
        self.boxes_mat_info = {}
        fir_keys = sorted(s_pl_bom.keys())
        q_dic = {}
        fp = {}
        for fir in fir_keys:
            sec_keys = sorted(s_pl_bom[fir])
            q_dic[fir] = {}
            for sec in sec_keys:
                th_key = sorted(s_pl_bom[fir][sec].keys())
                qty = 0.0
                qty1 = 0.0
                qty2 = 0.0
                if sec == '05002':
                    for th in th_key:
                        if '悬挂' in s_pl_bom[fir][sec][th]['mat_name']:
                            qty += s_pl_bom[fir][sec][th]['qty']

                            if '35*50' in s_pl_bom[fir][sec][th]['mat_name']:
                                fp[fir] = False
                            else:
                                fp[fir] = True

                    q_dic[fir][2] = qty

                elif sec == '05001':
                    for th in th_key:
                        if '门板' in s_pl_bom[fir][sec][th]['mat_name']:
                            qty1 += s_pl_bom[fir][sec][th]['qty']
                        elif '左门' in s_pl_bom[fir][sec][th]['mat_name']:
                            qty1 += s_pl_bom[fir][sec][th]['qty']

                        if '厅门装置' in s_pl_bom[fir][sec][th]['mat_name']:
                            qty2 += s_pl_bom[fir][sec][th]['qty']

                    q_dic[fir][1] = qty1
                    q_dic[fir][3] = qty2

        self.get_box_mat(fp, s_pl_bom, q_dic)

        return q_dic

    def get_box_mat(self, fp, s_pl_bom, q_dic):
        # 0 无法确认箱子信息， 1 - 只有打包件 厅门装置， 2 获得包装物料和数量
        cats = self.door_type_group.keys()

        for c in cats:
            self.boxes_mat_info[c] = {}
            d_keys = sorted(s_pl_bom[c])
            door_type = self.door_type_group[c][0][0]
            open_mode = self.door_type_group[c][0][1]
            dw = self.door_type_group[c][0][2]
            dh = self.door_type_group[c][0][3]
            em = self.door_type_group[c][0][4]
            if open_mode == '2 Panel Centre Open':
                om = 'CO'
            elif open_mode == '2 Panel Left Side Open' or open_mode == '2 Panel Right Side Open':
                om = '2S'
            elif open_mode == '2 Panel Centre Open and Fold':
                om = '2CO'

            for k in d_keys:
                mat_list = {}
                if k == '05001':
                    boxes = box_mat_logic.select().where((box_mat_logic.box_id == k) & (box_mat_logic.door_type == door_type) & (box_mat_logic.fire_resist == fp[c]) &
                                                         (box_mat_logic.open_type == om) & (box_mat_logic.elevator_type == em) &
                                                         (box_mat_logic.door_width_min <= dw) & (box_mat_logic.door_width_max >= dw) &
                                                         (box_mat_logic.door_height_min <= dh) & (box_mat_logic.door_height_max >= dh)).order_by(box_mat_logic.packing_qty_max.desc())
                elif k == '05002':
                    boxes = box_mat_logic.select().where((box_mat_logic.box_id == k) & (box_mat_logic.door_type == door_type) &
                                                         (box_mat_logic.open_type == om) & (box_mat_logic.elevator_type == em) &
                                                         (box_mat_logic.door_width_min <= dw) & (box_mat_logic.door_width_max >= dw)).order_by(box_mat_logic.packing_qty_max.desc())
                elif k == '05003' and door_type == 'S200':
                    if fp[c]:
                        jb = '35*50'
                    else:
                        jb = '35*50'
                    if int(dw) <= 900:
                        boxes = box_mat_logic.select().where((box_mat_logic.box_id == k) & (box_mat_logic.door_type == door_type) & (box_mat_logic.jamb_size == jb) &
                                                             (box_mat_logic.open_type == om) & (box_mat_logic.elevator_type == em) &
                                                             (box_mat_logic.door_width_min <= dw) & (box_mat_logic.door_width_max >= dw)).order_by(box_mat_logic.packing_qty_max.desc())
                    else:
                        boxes = box_mat_logic.select().where((box_mat_logic.box_id == k) & (box_mat_logic.door_type == door_type) &
                                                             (box_mat_logic.open_type == om) & (box_mat_logic.elevator_type == em) &
                                                             (box_mat_logic.door_width_min <= dw) & (box_mat_logic.door_width_max >= dw)).order_by(box_mat_logic.packing_qty_max.desc())

                if not boxes:
                    return 0

                i_s = 0
                for b in boxes:
                    i_s += 1
                    m_info = []

                    m_info.append(b.mat_no)
                    m_info.append(b.packing_qty_min)
                    m_info.append(b.packing_qty_max)
                    m_info.append(0)

                    mat_list[i_s] = m_info

                self.get_boxes_mats_info(c, k, mat_list, q_dic[c][2])
        return 2

    def get_boxes_mats_info(self, level, box_id, mat_list, s_qty):
        self.boxes_mat_info[level][box_id] = {}

        l_qty = s_qty
        i_pos = len(mat_list)

        i = 0
        for n in range(1, i_pos + 1):
            if mat_list[n][2] <= l_qty:
                i += 1
                q = int(l_qty / mat_list[n][2])
                l_qty = l_qty - q * mat_list[n][2]
                self.boxes_mat_info[level][box_id][i] = mat_list[n]
                if mat_list[n][1] <= l_qty:
                    self.boxes_mat_info[level][box_id][i][3] = q + 1
                    break
                else:
                    self.boxes_mat_info[level][box_id][i][3] = q

            if mat_list[n][2] > l_qty and l_qty >= mat_list[n][1]:
                i += 1
                self.boxes_mat_info[level][box_id][i] = mat_list[n]
                self.boxes_mat_info[level][box_id][i][3] = 1
                break

    def calc_pl(self):
        if len(self.wbs_bom) == 0:
            logger.info("请先导入或读取WBS BOM后在进行并箱操作!")
            return

        self.im_method = 2
        if self.data_thread is not None and self.data_thread.is_alive():
            messagebox.showinfo('提示', '表单刷新线程正在后台刷新列表，请等待完成后再点击!')
            return

        self.data_thread = refresh_thread(self)
        self.data_thread.setDaemon(True)
        self.data_thread.start()

    def get_unit_info_from_sap(self):
        self.prj_info_st = {}
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
        if l_wbs == 0:
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
                    self.prj_para_st[format_wbs_no(
                        re['POSID'])][re['ATNAM']] = re['ATWTB']
            else:
                self.prj_para_st[format_wbs_no(re['POSID'])] = {}
                self.prj_para_st[format_wbs_no(
                    re['POSID'])][re['ATNAM']] = re['ATWTB']

        logger.info("项目参数信息获取完成。")

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
            if re['VORNR'] not in merge_op:
                continue

            i += 1
            line = {}
            wbs = re['POSID']

            re['PSPEL'] = act_to_wbs_element(wbs, re['VORNR'])

            line['mat_no'] = re['MATNR_I']
            line['mat_name'] = re['MAKTX_ZH']
            line['wbs_element'] = re['PSPEL']
            line['qty'] = float(re['NOMNG'])
            line['wbs_no'] = wbs
            line['remarks'] = str(re['POTX1']).strip()
            line['box_id'] = ''
            line['is_box_mat'] = False
            line['activity'] = re['VORNR']

            self.wbs_bom[i] = line

    def read_packing_mode(self, s_type, b_wooden=True):
        re = door_packing_mode.select().where((door_packing_mode.door_type == s_type) & (door_packing_mode.is_wooden == b_wooden))\
            .order_by(door_packing_mode.box_id.asc())

        if not re:
            return None

        packing_mode = {}
        for r in re:
            s_s = {}
            if r.packing_desc is not None:
                s_s[1] = r.packing_desc.split(';')

            if r.packing_except is not None:
                s_s[2] = r.packing_except.split(';')

            packing_mode[r.box_id] = s_s

        return packing_mode

    def group_door_type(self):
        if len(self.prj_para_st) == 0:
            return

        i = 0
        keys = sorted(self.prj_para_st.keys())
        for key in keys:
            d1 = self.prj_para_st[key]['TC041']
            d2 = self.prj_para_st[key]['TC036']
            d3 = self.prj_para_st[key]['TC037']
            d4 = self.prj_para_st[key]['TC038']
            d5 = self.prj_para_st[key]['TC001'][:2]

            ld = [d1, d2, d3, d4, d5]

            k = self.check_same_door_type(ld)

            if k >= 0:
                self.door_type_group[k][1].append(key)
            else:
                i += 1
                self.door_type_group[i] = [ld, [key]]

    def check_same_door_type(self, l):
        for k in self.door_type_group.keys():
            if self.door_type_group[k][0] == l:
                return k

        return -1

    def scan_bom_fill_boxid(self):
        self.packing_bom = {}
        old_box_list = {}

        i_len = len(self.wbs_bom)

        if i_len == 0:
            return

        wbs = ''

        for i in range(1, i_len + 1):
            b_add = True
            line = self.wbs_bom[i].copy()
            if wbs != line['wbs_no']:
                wbs = line['wbs_no']
                d_type = self.prj_para_st[wbs]['TC041']

                if 'S8' in d_type.upper():
                    p_mode = self.read_packing_mode('S8')
                else:
                    b_wood = self.is_wooden.get()
                    p_mode = self.read_packing_mode(d_type, b_wood)

            if '包装箱' in line['mat_name']:
                i_pos = line['mat_name'].find('5#')
                b_id = line['mat_name'][i_pos:i_pos + 4]
                box_id = box_to_id(b_id)
                line['is_box_mat'] = True
            else:
                box_id = self.get_boxid_in_bom(line, p_mode)

            self.wbs_bom[i]['box_id'] = box_id
            line['box_id'] = box_id

            if wbs not in self.packing_bom.keys():
                self.packing_bom[wbs] = {}

            if not box_id in self.packing_bom[wbs].keys():
                self.packing_bom[wbs][box_id] = {}

            j = len(self.packing_bom[wbs][box_id])

            for m in range(1, j + 1):
                if line['mat_no'] == self.packing_bom[wbs][box_id][m]['mat_no'] and \
                        line['remarks'] == self.packing_bom[wbs][box_id][m]['remarks']:
                    self.packing_bom[wbs][box_id][m]['qty'] += line['qty']
                    b_add = False
                    break

            m = len(old_box_list)
            if b_add and line['is_box_mat'] == False:
                self.packing_bom[wbs][box_id][j + 1] = line
            elif line['is_box_mat'] == True:
                old_box_list[m + 1] = line

    def get_boxid_in_bom(self, line, p_mode):
        if not p_mode:
            return '05000'

        for key in p_mode.keys():
            if len(p_mode[key]) == 0:
                default = key

        m_name = line['mat_name']
        for key in p_mode.keys():
            if len(p_mode[key]) == 0:
                continue

            for p in p_mode[key][1]:
                if p in m_name:
                    if 2 in p_mode[key].keys():
                        for q in p_mode[key][2]:
                            if q in m_name:
                                return default
                    return key

        return default

    def sum_mat_qty(self):
        i_len = len(self.door_type_group)

        if i_len == 0:
            return None

        sum_packing_bom = {}
        for i in range(1, i_len + 1):
            li_wbs = self.door_type_group[i][1]

            sum_packing_bom[i] = self.sum_qty(li_wbs)

        return sum_packing_bom

    def sum_qty(self, li):
        sum_bm = {}

        for l in li:
            if len(sum_bm) == 0:
                sum_bm = copy.deepcopy(self.packing_bom[l])
                continue

            keys = sorted(sum_bm.keys())
            n_bm = copy.deepcopy(self.packing_bom[l])

            for key in keys:
                for n in list(n_bm[key]):
                    m_len = len(sum_bm[key])
                    for m in list(sum_bm[key]):
                        if sum_bm[key][m]['mat_no'] == n_bm[key][n]['mat_no'] and \
                                sum_bm[key][m]['remarks'] == n_bm[key][n]['remarks']:
                            sum_bm[key][m]['qty'] += n_bm[key][n]['qty']
                            break
                        elif m == m_len:
                            sum_bm[key][m_len +
                                        1] = copy.deepcopy(n_bm[key][n])
        return sum_bm
