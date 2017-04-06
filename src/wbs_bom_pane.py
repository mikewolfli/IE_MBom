# coding=utf-8
'''
Created on 2017年3月17日

@author: 10256603
'''
from global_list import *
global login_info


logger = logging.getLogger()

wbs_headers = (('POSID', 'Work Breakdown Structure Element (WBS Element)'),
               ('MATNR_H', 'Material Number'),
               ('WERKS', 'Plant'),
               ('STLAN', 'BOM Usage'),
               ('POSNR', 'Old: Project number : No longer used --> PS_POSNR'),
               ('POSTP', 'Item Category (Bill of Material)'),
               ('MATNR_I', 'Material Number'),
               ('MAKTX_ZH', 'Material Description (Short Text)'),
               ('MAKTX_EN', 'Material Description (Short Text)'),
               ('PSPEL', 'WBS Element'),
               ('VORNR', 'Operation/Activity Number'),
               ('BDTER', 'Requirement Date for the Component'),
               ('NOMNG', 'Required quantity'),
               ('ENMNG', 'Quantity Withdrawn'),
               ('PRLAB', 'Valuated Unrestricted-Use Stock'),
               ('KZEAR', 'Final Issue for This Reservation'),
               ('POTX2', 'BOM item text (line 2)'),
               ('POTX1', 'BOM Item Text (Line 1)'),
               )


prj_info_header = (('PSPID', 'Project Definition'),
                   ('POST1', 'PS: Short description (1st text line)'),
                   ('POSID', 'Work Breakdown Structure Element (WBS Element)'),
                   ('VBELN', 'Sales Document'),
                   ('POSNR', 'Sales Document Item'),
                   ('KUNNR', 'Sold-to party'),
                   ('ORT01', 'City'),
                   ('NAME1', 'Name 1'),
                   ('LAND1', 'Country Key'),
                   ('BSTKD', 'Purchase order number'),
                   ('USR01', '2nd user field 20 digits - WBS element'),
                   ('VGBEL', 'Contract number'),
                   ('CONTSIGND', 'Contract Signed Date'),
                   ('CONRECD', 'Contract Recieved Date'),
                   ('CONLCHGD_CH', 'Contract changed date'),
                   ('CONFEEBKRELD', 'Contract Problem Feeback Sheet Release Date'),
                   ('CONFEEBKREPD', 'Contract Problem feedback sheet Reply Date'),
                   ('CONREQDELD', 'Contract Required Delivery Date (leadtime)'),
                   ('CONLCHGD', 'Date of Last Changed Contract'),
                   ('GADRECD', 'Gad Recieve Date'),
                   ('LGADCHGD', 'Last Gad change date'),
                   ('SCRUFIND', 'GAD Confirm date(field)'),
                   ('RECSTNOTID', 'Project Start Notice Recieve Date'),
                   ('PRJSTD', 'Project Start Date(internal)'),
                   ('SPECMEMORECD', 'Spec Memo Recieve Date'),
                   ('ZTRANFERDS', 'Tranfer to DS Date'),
                   ('SPECMEMOD', 'Spec Memo Confirm Date'),
                   ('RECSHPD', 'Received date of shipping notice'),
                   ('REPSHPD', 'Repy date for shipping noticed'),
                   ('TPLISTD', 'Full delivery date'),
                   ('TPCONFD', 'TP confirmed date'),
                   ('TURNOVERPD', 'tun over date (plan)'),
                   ('ASSFINID', 'Assessment finish date'),
                   ('PROJFLAG', 'Project start/change Notice'),
                   ('REPTNOTID', 'Reply Date for Project Start Notice'),
                   ('WAFIND', 'Warranty finish date'),
                   ('CONFFINPD', 'Configration Finish Date(plan)'),
                   ('CONFFIND', 'Configration Finish Date'),
                   ('ZACCDDATE', 'Actual Cash Collection Date'),
                   ('ZSADATE', 'Special Approval Date'),
                   ('ZCHGDATE', 'Change on date'),
                   ('ZCHGBY', 'Change by'),
                   ('ZFIELDWBS', 'Field WBS'),
                   ('BATCH1', 'unit batch1'),
                   ('BATCH2', 'unit batch2'),
                   ('BATCH3', 'unit batch3'),
                   ('BATCH4', 'unit batch3'),
                   ('BATCH5', 'unit batch3'),
                   ('BATCH6', 'unit batch3'),
                   ('BATCH7', 'unit batch3'),
                   ('BATCHDESI1', 'Batch Description input'),
                   ('BATCHDESI2', 'Batch Description input'),
                   ('BATCHDESI3', 'Batch Description input'),
                   ('BATCHDESI4', 'Batch Description input'),
                   ('BATCHDESI5', 'Batch Description input'),
                   ('BATCHDESI6', 'Batch Description input'),
                   ('BATCHDESI7', 'Batch Description input'),
                   ('BATCHDESO1', 'Batch Description output'),
                   ('BATCHDESO2', 'Batch Description output'),
                   ('BATCHDESO3', 'Batch Description output'),
                   ('BATCHDESO4', 'Batch Description output'),
                   ('BATCHDESO5', 'Batch Description output'),
                   ('BATCHDESO6', 'Batch Description output'),
                   ('BATCHDESO7', 'Batch Description output'),
                   ('ZBATCHDESI1', 'Batch Description input extend'),
                   ('ZBATCHDESI2', 'Batch Description input extend'),
                   ('ZBATCHDESI3', 'Batch Description input extend'),
                   ('ZBATCHDESI4', 'Batch Description input extend'),
                   ('ZBATCHDESI5', 'Batch Description input extend'),
                   ('ZBATCHDESI6', 'Batch Description input extend'),
                   ('ZBATCHDESI7', 'Batch Description input extend'),
                   ('BATCHCOMBIN1', 'Batch Combination'),
                   ('BATCHCOMBIN2', 'Batch Combination'),
                   ('BATCHCOMBIN3', 'Batch Combination'),
                   ('BATCHCOMBIN4', 'Batch Combination'),
                   ('BATCHCOMBIN5', 'Batch Combination'),
                   ('BATCHCOMBIN6', 'Batch Combination'),
                   ('BATCHCOMBIN7', 'Batch Combination'),
                   ('PLDELD1', 'Planned Delivery Date after Neogotiation'),
                   ('PLDELD2', 'Planned Delivery Date after Neogotiation'),
                   ('PLDELD3', 'Planned Delivery Date after Neogotiation'),
                   ('PLDELD4', 'Planned Delivery Date after Neogotiation'),
                   ('PLDELD5', 'Planned Delivery Date after Neogotiation'),
                   ('PLDELD6', 'Planned Delivery Date after Neogotiation'),
                   ('PLDELD7', 'Planned Delivery Date after Neogotiation'),
                   ('ZSTODES1', 'Batch Description output extend'),
                   ('ZSTODES2', 'Batch Description output extend'),
                   ('ZSTODES3', 'Batch Description output extend'),
                   ('ZSTODES4', 'Batch Description output extend'),
                   ('ZSTODES5', 'Batch Description output extend'),
                   ('ZSTODES6', 'Batch Description output extend'),
                   ('ZSTODES7', 'Batch Description output extend'),
                   ('REQSHPD1', 'Request delivery date by shipping notice'),
                   ('REQSHPD2', 'Request delivery date by shipping notice'),
                   ('REQSHPD3', 'Request delivery date by shipping notice'),
                   ('REQSHPD4', 'Request delivery date by shipping notice'),
                   ('REQSHPD5', 'Request delivery date by shipping notice'),
                   ('REQSHPD6', 'Request delivery date by shipping notice'),
                   ('REQSHPD7', 'Request delivery date by shipping notice'),
                   ('CONSHPD1', 'Confirmed date for delivery date for shipping notice'),
                   ('CONSHPD2', 'Confirmed date for delivery date for shipping notice'),
                   ('CONSHPD3', 'Confirmed date for delivery date for shipping notice'),
                   ('CONSHPD4', 'Confirmed date for delivery date for shipping notice'),
                   ('CONSHPD5', 'Confirmed date for delivery date for shipping notice'),
                   ('CONSHPD6', 'Confirmed date for delivery date for shipping notice'),
                   ('CONSHPD7', 'Confirmed date for delivery date for shipping notice'),
                   ('ZBATCHDESO1', 'Batch Description output extend'),
                   ('ZBATCHDESO2', 'Batch Description output extend'),
                   ('ZBATCHDESO3', 'Batch Description output extend'),
                   ('ZBATCHDESO4', 'Batch Description output extend'),
                   ('ZBATCHDESO5', 'Batch Description output extend'),
                   ('ZBATCHDESO6', 'Batch Description output extend'),
                   ('ZBATCHDESO7', 'Batch Description output extend'),
                   ('PURFIND1', 'Purchase Finish Date'),
                   ('PURFIND2', 'Purchase Finish Date'),
                   ('PURFIND3', 'Purchase Finish Date'),
                   ('PURFIND4', 'Purchase Finish Date'),
                   ('PURFIND5', 'Purchase Finish Date'),
                   ('PURFIND6', 'Purchase Finish Date'),
                   ('PURFIND7', 'Purchase Finish Date'),
                   ('ZFINDELD1', 'Finalize delivery date'),
                   ('ZFINDELD2', 'Finalize delivery date'),
                   ('ZFINDELD3', 'Finalize delivery date'),
                   ('ZFINDELD4', 'Finalize delivery date'),
                   ('ZFINDELD5', 'Finalize delivery date'),
                   ('ZFINDELD6', 'Finalize delivery date'),
                   ('ZFINDELD7', 'Finalize delivery date'),
                   ('DELSTNOTID1', 'Delivery Date Required by Project Start Notice'),
                   ('DELSTNOTID2', 'Delivery Date Required by Project Start Notice'),
                   ('DELSTNOTID4', 'Delivery Date Required by Project Start Notice'),
                   ('DELSTNOTID5', 'Delivery Date Required by Project Start Notice'),
                   ('DELSTNOTID6', 'Delivery Date Required by Project Start Notice'),
                   ('DELSTNOTID7', 'Delivery Date Required by Project Start Notice'),
                   ('ZSTOPRED1', 'Stock Preparation date'),
                   ('ZSTOPRED2', 'Stock Preparation date'),
                   ('ZSTOPRED3', 'Stock Preparation date'),
                   ('ZSTOPRED4', 'Stock Preparation date'),
                   ('ZSTOPRED5', 'Stock Preparation date'),
                   ('ZSTOPRED6', 'Stock Preparation date'),
                   ('ZSTOPRED7', 'Stock Preparation date'),
                   ('Z_AP_PROJCLAS', 'Project classification'),
                   ('Z_AP_PROJTYPE', 'project type'),
                   ('ZPLGR1', 'Condition type YCP1(Basic Price)'),
                   ('CHK1', 'GAD FLAG1:Confirmed GAD'),
                   ('CHK2', 'GAD FLAG2 : PSN'),
                   ('CHK3', 'GAD FLAG3: design'),
                   ('CHK4', 'Latest GAD'),
                   ('VKAUS', 'Usage Indicator'),
                   ('AUGRU', 'Order reason (reason for the business transaction)'),
                   ('AUDAT', 'Document Date (Date Received/Sent)'),
                   ('AUART', 'Sales Document Type'),
                   ('MATNR', 'Material Number'),
                   ('RFBSK', 'Status for transfer to accounting'),
                   ('FKDAT', 'Billing date for billing index and printout'),
                   ('STAT', 'Project status'),
                   )

dict_act_box = {'0010': '.02',
                '0020': '.02',
                '0030': '.02',
                '0040': '.02',
                '0050': '.02',
                '0060': '.02',
                '0070': '.02',
                '0080': '.02',
                '0090': '.02',
                '0100': '.02',
                '0110': '.02',
                '0120': '.05',
                '0130': '.05',
                '0140': '.08',
                '0150': '.05',
                '0160': '.12',
                '0170': '.12',
                '0180': '.12',
                '0190': '.05',
                '0200': '.05',
                '0210': '.08',
                '0220': '.05',
                '0230': '.06',
                '0240': '.03',
                '0250': '.06',
                '0260': '.09',
                '0270': '.03',
                '0280': '.06',
                '0290': '.09',
                '0300': '.03',
                '0310': '.15',
                '0320': '.09',
                '0330': '.10',
                '0340': '.01',
                '0350': '.09',
                '0360': '.13',
                '0370': '.13',
                '0380': '.09',
                '0390': '.12',
                '0400': '.08',
                '0410': '.03',
                '0420': '.10',
                '0430': '.08',
                '0440': '.03',
                '0450': '.16',
                '0460': '.06',
                '0470': '.15',
                '0480': '.03',
                '0490': '.09',
                '0500': '.10',
                '0510': '.08',
                '0520': '.05',
                '0530': '.08',
                '0540': '.11',
                '0550': '.11',
                '0560': '.12',
                '0570': '.08',
                '0580': '.11',
                '0590': '.11',
                '0600': '.12',
                '0610': '.11',
                '0620': '.03',
                '0630': '.01',
                '0640': '.01',
                '0650': '.01',
                '0660': '.07',
                '0670': '.08',
                '0680': '.04',
                '0690': '.03',
                '0700': '.04',
                '0710': '.08',
                '0720': '.09',
                '0730': '.03',
                '0740': '.13',
                '0750': '.15',
                '0760': '.14',
                '0770': '.08',
                '0780': '.08',
                '0790': '.02',
                '0800': '.17',
                }
'''
prj_para_header = (('POSID','Work Breakdown Structure Element (WBS Element)'),
                   ('ATNAM','Characteristic Name'),
                   ('ATBEZ','Characteristic description'),
                   ('ATWRT','Characteristic Value'),
                   ('ATWTB','Characteristic value description'),
                   )
'''

def act_to_wbs_element(wbs, act):
    if act not in dict_act_box.keys():
        return ''
        
    return wbs+dict_act_box[act]

def dict_has_key(di, key):
    if key in di.keys():
        return True
    else:
        return False
    
def format_wbs_no(wbs):
    if len(wbs)==0:
        return ''
    
    return wbs[0]+'/'+wbs[1:9]+'.'+wbs[9:12]
    
class wbs_bom_pane(Frame):
    data_thread = None
    wbses = []
    mates = []
    prj_para_header = [
        ('POSID', 'Work Breakdown Structure Element (WBS Element)'), ]

    wbs_bom = {}
    prj_info_st = {}
    prj_para_st = {}

    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.grid()

        self.createWidgets()

    def createWidgets(self):
        self.imput_group = Frame(self)
        self.imput_group.grid(row=0, column=0, rowspan=3,
                              columnspan=5, sticky=NSEW)

        self.wbs_label = Label(self.imput_group, text='WBS NO.', anchor='w')
        self.wbs_label.grid(row=0, column=0, sticky=EW)

        self.mat_label = Label(self.imput_group, text='物料号', anchor='w')
        self.mat_label.grid(row=1, column=0, sticky=EW)

        self.plant_label = Label(self.imput_group, text='工厂', anchor='w')
        self.plant_label.grid(row=2, column=0, sticky=EW)

        self.wbs_str = StringVar()
        self.wbs_entry = Entry(self.imput_group, textvariable=self.wbs_str)
        self.wbs_entry.grid(row=0, column=1, columnspan=3, sticky=EW)
        self.wbs_entry['state'] = 'readonly'

        self.mat_str = StringVar()
        self.mat_entry = Entry(self.imput_group, textvariable=self.mat_str)
        self.mat_entry.grid(row=1, column=1, columnspan=3, sticky=EW)
        self.mat_entry['state'] = 'readonly'

        self.plant_str = StringVar()
        self.plant_entry = Entry(self.imput_group, textvariable=self.plant_str)
        self.plant_entry.grid(row=2, column=1, columnspan=1, sticky=EW)
        self.plant_str.set('2101')

        self.plant_desc1 = Label(self.imput_group, text='(默认)', anchor='w')
        self.plant_desc1.grid(row=2, column=2, sticky=EW)

        self.plant_desc2 = Label(
            self.imput_group, text='中山工厂为2001', anchor='w')
        self.plant_desc2.grid(row=2, column=3, sticky=EW)

        self.wbs_button = Button(self.imput_group, text='...')
        self.wbs_button.grid(row=0, column=4, sticky=EW)
        self.wbs_button['command'] = self.get_wbs_list

        self.mat_button = Button(self.imput_group, text='...')
        self.mat_button.grid(row=1, column=4, sticky=EW)
        self.mat_button['command'] = self.get_mat_list

        self.search_button = Button(self, text='查询')
        self.search_button.grid(row=0, column=5, sticky=EW)
        self.search_button['command'] = self.start_search

        self.clear_button = Button(self, text='清空条件')
        self.clear_button.grid(row=1, column=5, sticky=EW)
        self.clear_button['command'] = self.clear_case

        self.export_result = Button(self, text='导出结果')
        self.export_result.grid(row=2, column=5, sticky=EW)
        self.export_result['command'] = self.export_res

        self.with_prj_info = BooleanVar()
        self.with_prj_para = BooleanVar()

        check_prj_info = Checkbutton(self, text="显示项目信息", variable=self.with_prj_info,
                                     onvalue=True, offvalue=False)
        check_prj_info.grid(row=0, column=6, sticky=EW)

        check_prj_para = Checkbutton(self, text="显示项目参数", variable=self.with_prj_para,
                                     onvalue=True, offvalue=False)

        check_prj_info['command'] = self.with_prj_info_check

        check_prj_para['command'] = self.with_prj_para_check

        check_prj_para.grid(row=1, column=6, sticky=EW)

        self.ntbook = ttk.Notebook(self)
        self.ntbook.rowconfigure(0, weight=1)
        self.ntbook.columnconfigure(0, weight=1)

        self.wbs_tab = Frame(self)

        style = ttk.Style()
        style.configure("Treeview", font=('TkDefaultFont', 10))
        style.configure("Treeview.Heading", font=('TkDefaultFont', 9))

        wbs_cols = ['col1', 'col2', 'col3', 'col4', 'col5', 'col6', 'col7', 'col8',
                    'col9', 'col10', 'col11', 'col12', 'col13', 'col14', 'col15', 'col16', 'col17',
                    'col18']
        self.wbs_list = ttk.Treeview(
            self.wbs_tab, show='headings', columns=wbs_cols, selectmode='browse')
        self.wbs_list.grid(row=0, column=0, rowspan=6,
                           columnspan=10, sticky='nsew')

        for i in range(0, len(wbs_cols)):
            self.wbs_list.heading(wbs_cols[i], text=wbs_headers[i][1])

        self.wbs_list.column('col1', width=80, anchor='w')
        self.wbs_list.column('col2', width=80, anchor='w')
        self.wbs_list.column('col3', width=40, anchor='w')
        self.wbs_list.column('col4', width=40, anchor='w')
        self.wbs_list.column('col5', width=40, anchor='w')
        self.wbs_list.column('col6', width=40, anchor='w')
        self.wbs_list.column('col7', width=80, anchor='w')
        self.wbs_list.column('col8', width=150, anchor='w')
        self.wbs_list.column('col9', width=150, anchor='w')
        self.wbs_list.column('col10', width=80, anchor='w')
        self.wbs_list.column('col11', width=60, anchor='w')
        self.wbs_list.column('col12', width=80, anchor='w')
        self.wbs_list.column('col13', width=60, anchor='w')
        self.wbs_list.column('col14', width=60, anchor='w')
        self.wbs_list.column('col15', width=40, anchor='w')
        self.wbs_list.column('col16', width=40, anchor='w')
        self.wbs_list.column('col17', width=200, anchor='w')
        self.wbs_list.column('col18', width=200, anchor='w')

        wbs_ysb = ttk.Scrollbar(self.wbs_tab, orient='vertical',
                                command=self.wbs_list.yview)
        wbs_xsb = ttk.Scrollbar(self.wbs_tab, orient='horizontal',
                                command=self.wbs_list.xview)
        wbs_ysb.grid(row=0, column=10, rowspan=6, sticky='ns')
        wbs_xsb.grid(row=6, column=0, columnspan=10, sticky='ew')

        self.wbs_list.configure(yscroll=wbs_ysb.set, xscroll=wbs_xsb.set)

        self.wbs_tab.rowconfigure(1, weight=1)
        self.wbs_tab.columnconfigure(1, weight=1)

        self.ntbook.add(self.wbs_tab, text='WBS BOM', sticky=NSEW)
        self.ntbook.grid(row=3, column=0, rowspan=12,
                         columnspan=10, sticky='nsew')

        self.prj_info_tab = Frame(self)

        i_range = len(prj_info_header)
        prj_info_cols = ['col' + str(i) for i in range(i_range)]

        self.prj_info_list = ttk.Treeview(
            self.prj_info_tab, show='headings', columns=prj_info_cols, selectmode='browse')

        self.prj_info_list.grid(row=0, column=0, rowspan=6,
                                columnspan=10, sticky='nsew')

        for i in range(0, i_range):
            self.prj_info_list.heading(
                prj_info_cols[i], text=prj_info_header[i][1])

        for i in range(0, i_range):
            self.prj_info_list.column(prj_info_cols[i], width=100, anchor='w')

        prj_info_ysb = ttk.Scrollbar(self.prj_info_tab, orient='vertical',
                                     command=self.prj_info_list.yview)
        prj_info_xsb = ttk.Scrollbar(self.prj_info_tab, orient='horizontal',
                                     command=self.prj_info_list.xview)
        prj_info_ysb.grid(row=0, column=10, rowspan=6, sticky='ns')
        prj_info_xsb.grid(row=6, column=0, columnspan=10, sticky='ew')

        self.prj_info_list.configure(
            yscroll=prj_info_ysb.set, xscroll=prj_info_xsb.set)

        self.prj_info_tab.rowconfigure(1, weight=1)
        self.prj_info_tab.columnconfigure(1, weight=1)

        self.ntbook.add(self.prj_info_tab, text='项目信息', sticky=NSEW)

        if not self.with_prj_info.get():
            self.ntbook.hide(self.prj_info_tab)

        self.get_prj_para_headers()
        self.prj_para_tab = Frame(self)

        i_range = len(self.prj_para_header)
        prj_para_cols = ['col' + str(i) for i in range(i_range)]

        self.prj_para_list = ttk.Treeview(
            self.prj_para_tab, show='headings', columns=prj_para_cols, selectmode='browse')

        self.prj_para_list.grid(row=0, column=0, rowspan=6,
                                columnspan=10, sticky='nsew')

        for i in range(0, i_range):
            self.prj_para_list.heading(
                prj_para_cols[i], text=self.prj_para_header[i][1])

        for i in range(0, i_range):
            self.prj_para_list.column(prj_para_cols[i], width=100, anchor='w')

        prj_para_ysb = ttk.Scrollbar(self.prj_para_tab, orient='vertical',
                                     command=self.prj_para_list.yview)
        prj_para_xsb = ttk.Scrollbar(self.prj_para_tab, orient='horizontal',
                                     command=self.prj_para_list.xview)
        prj_para_ysb.grid(row=0, column=10, rowspan=6, sticky='ns')
        prj_para_xsb.grid(row=6, column=0, columnspan=10, sticky='ew')

        self.prj_para_list.configure(
            yscroll=prj_para_ysb.set, xscroll=prj_para_xsb.set)

        self.prj_para_tab.rowconfigure(1, weight=1)
        self.prj_para_tab.columnconfigure(1, weight=1)

        self.ntbook.add(self.prj_para_tab, text='项目参数', sticky=NSEW)

        if not self.with_prj_para.get():
            self.ntbook.hide(self.prj_para_tab)

        log_pane = Frame(self)

        self.log_label = Label(log_pane)
        self.log_label["text"] = "操作记录"
        self.log_label.grid(row=0, column=0, sticky=W)

        self.log_text = scrolledtext.ScrolledText(log_pane, state='disabled')
        self.log_text.config(font=('TkFixedFont', 10, 'normal'))
        self.log_text.grid(row=1, column=0, columnspan=2, sticky=EW)
        log_pane.rowconfigure(1, weight=1)
        log_pane.columnconfigure(1, weight=1)

        log_pane.grid(row=15, column=0, rowspan=2, columnspan=10, sticky=NSEW)

        # Create textLogger
        text_handler = TextHandler(self.log_text)
        # Add the handler to logger

        logger.addHandler(text_handler)
        logger.setLevel(logging.INFO)

        self.rowconfigure(6, weight=1)
        self.columnconfigure(9, weight=1)

    def get_wbs_list(self):
        d = ask_list("WBS NO输入", 3)
        if not d:
            return

        self.wbses = d
        self.wbs_str.set(self.wbses[0])

    def get_mat_list(self):
        d = ask_list("物料输入", 2)
        if not d:
            return

        self.mates = d
        self.mat_str.set(self.mates[0])

    def start_search(self):
        if len(self.wbses) == 0 and len(self.mates) == 0:
            return

        if self.data_thread is not None and self.data_thread.is_alive():
            messagebox.showinfo('提示', '表单刷新线程正在后台刷新列表，请等待完成后再点击!')
            return

        self.__refresh_tree()

    def clear_case(self):
        self.mat_str.set('')
        self.wbs_str.set('')
        self.wbses = []
        self.mates = []
        self.wbs_bom = {}
        self.prj_info_st = {}
        self.prj_para_st = {}
        self.plant_str.set('2101')

        self.clear_list()

    def clear_list(self):
        logger.info("正在清空列表...")
        for row in self.wbs_list.get_children():
            self.wbs_list.delete(row)

        if self.with_prj_info.get():
            for row in self.prj_info_list.get_children():
                self.prj_info_list.delete(row)

        if self.with_prj_para.get():
            for row in self.prj_para_list.get_children():
                self.prj_para_list.delete(row)

    def export_res(self):
        if len(self.wbs_bom)==0:
            logger.info("查询结果为空，无法导出清单")
            return
        
        file_str = filedialog.asksaveasfilename(
            title="导出文件", filetypes=[('excel file', '.xlsx')])
        if not file_str:
            return

        if not file_str.endswith(".xlsx"):
            file_str += ".xlsx"

        wb = Workbook()
        ws = wb.worksheets[0]
        ws.title = 'WBS BOM清单'
        col_size = len(wbs_headers)
        for i in range(0, col_size):
            ws.cell(row=1, column=i + 1).value = wbs_headers[i][1]

        i_wbs_bom = len(self.wbs_bom)

        for i in range(1, i_wbs_bom + 1):
            for j in range(0, col_size):
                ws.cell(row=i + 1, column=j + 1).value = self.wbs_bom[i][j]
                
        if self.with_prj_info.get():
            ws1=wb.create_sheet("项目基本信息")
            
            for i in range(0, len(prj_info_header)):
                ws1.cell(row=1, column=i + 1).value = prj_info_header[i][1]
                
            i_prj_info = len(self.prj_info_st)
            for i in range(1, i_prj_info+1):
                for j in range(0, len(prj_info_header)):
                    ws1.cell(row=i + 1, column=j + 1).value = self.prj_info_st[i][j]
                    
        if self.with_prj_para.get():
            ws2 = wb.create_sheet("项目参数信息")
            for i in range(0, len(self.prj_para_header)):
                ws2.cell(row=1, column=i+1).value = self.prj_para_header[i][1]
                
            i_prj_para = len(self.prj_para_st)
            for i in range(1, i_prj_para+1):
                for j in range(0, len(self.prj_para_header)):
                    ws2.cell(row=i + 1, column=j + 1).value = self.prj_para_st[i][j]                   
            
        if excel_xlsx.save_workbook(workbook=wb, filename=file_str):
            messagebox.showinfo("输出", "成功输出!")

    def with_prj_info_check(self):
        b_prj_info = self.with_prj_info.get()

        if b_prj_info:
            self.ntbook.add(self.prj_info_tab)
        else:
            self.ntbook.hide(self.prj_info_tab)

    def with_prj_para_check(self):
        b_prj_para = self.with_prj_para.get()

        if b_prj_para:
            self.ntbook.add(self.prj_para_tab)
        else:
            self.ntbook.hide(self.prj_para_tab)

    def get_prj_para_headers(self):
        re = SParameterFields.select().order_by(SParameterFields.field_name.asc())

        if not re:
            self.with_prj_para.set(False)
            return

        for r in re:
            li_1 = r.field_name
            li_2 = r.field_desc
            li = (li_1, li_2)
            self.prj_para_header.append(li)

    def refresh(self):
        self.wbs_bom = {}
        self.prj_info_st = {}
        self.prj_para_st = {}

        self.clear_list()

        logger.info("正在登陆SAP...")
        config = ConfigParser()
        config.read('sapnwrfc.cfg')
        para_conn = config._sections['connection']
        para_conn['user'] = base64.b64decode(para_conn['user']).decode()
        para_conn['passwd'] = base64.b64decode(para_conn['passwd']).decode()


        try:
            logger.info("正在连接SAP...")
            conn = pyrfc.Connection(**para_conn)

            wbs = self.refresh_wbs_bom(conn)

            if wbs is None:
                return 0

            self.refresh_prj(conn, wbs)

            self.update_tree()

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

        l_wbs = len(self.wbses)
        l_mat = len(self.mates)

        if l_wbs > 0 and l_mat > 0:
            for w in self.wbses:
                for m in self.mates:
                    line = dict(POSID=w, MATNR=m)
                    imp.append(line)
        elif l_wbs == 0 and l_mat > 0:
            for m in self.mates:
                line = dict(MATNR=m)
                imp.append(line)
        elif l_wbs > 0 and l_mat == 0:
            for w in self.wbses:
                line = dict(POSID=w)
                imp.append(line)

        else:
            return None

        logger.info("正在调用RFC函数(ZAP_PS_WBSBOM_INFO)...")
        result = conn.call('ZAP_PS_WBSBOM_INFO',
                           IT_CE_WBSBOM=imp, CE_WERKS=self.plant_str.get())

        i = 0
        wbs_s = []
        logger.info("正在分析并获取WBS BOM清单...")
        for re in result['OT_WBSBOM']:
            i += 1
            r_line = []
            wbs = re['POSID']
            if wbs not in wbs_s:
                wbs_s.append(wbs)
            for w in wbs_headers:
                if w[0] =='PSPEL':
                    r_line.append(act_to_wbs_element(wbs, re['VORNR']))
                else:
                    r_line.append(re[w[0]])
            self.wbs_bom[i] = r_line
        
        logger.info("wbs bom清单获取完成。")
        return wbs_s

    def dict_to_list(self, dict_para):
        i = 0
        i_len = len(self.prj_para_header)
        for key in dict_para.keys():
            i += 1
            s_line = dict_para[key]
            a_line = []
            a_line.append(key)

            for j in range(1, i_len):
                if dict_has_key(s_line, self.prj_para_header[j][0]):
                    a_line.append(s_line[self.prj_para_header[j][0]])
                else:
                    a_line.append('')

            self.prj_para_st[i] = a_line            

    def refresh_prj(self, conn, wbs_s):
        imp = []

        l_wbs = len(wbs_s)

        if l_wbs > 0:
            for w in wbs_s:
                line = dict(POSID=w)
                imp.append(line)
        elif l_wbs == 0:
            return

        b_prj_info = self.with_prj_info.get()
        b_prj_para = self.with_prj_info.get()

        if b_prj_info or b_prj_para:
            logger.info("正在调用RFC函数(ZAP_PS_PROJECT_INFO)...")
            result = conn.call('ZAP_PS_PROJECT_INFO',
                               IT_CE_POSID=imp, CE_WERKS=self.plant_str.get())

            if b_prj_info:
                logger.info("正在分析并获取项目基本信息...")

                i = 0
                for re in result['OT_PROJ']:
                    i += 1
                    r_line = []
                                    
                    for w in prj_info_header:
                        if w[0] == 'POSID':
                            r_line.append(format_wbs_no(re[w[0]]))
                        else:
                            r_line.append(re[w[0]])

                    self.prj_info_st[i] = r_line
                logger.info("项目基本信息获取完成")

            if b_prj_para:
                logger.info("正在分析并获取项目参数信息...")

                prj_para_s = {}
                
                for re in result['OT_CONF']:
                    
                    if dict_has_key(prj_para_s, format_wbs_no(re['POSID'])):
                        if not dict_has_key(prj_para_s[format_wbs_no(re['POSID'])], re['ATNAM']):
                            prj_para_s[format_wbs_no(re['POSID'])][re['ATNAM']] = re['ATWTB']
                    else:
                        prj_para_s[format_wbs_no(re['POSID'])] = {}
                        prj_para_s[format_wbs_no(re['POSID'])][re['ATNAM']] = re['ATWTB']

                if len(prj_para_s) != 0:
                    self.dict_to_list(prj_para_s)
                    
                logger.info("项目参数信息获取完成。")

    def __refresh_tree(self):
        self.data_thread = refresh_thread(self)
        self.data_thread.setDaemon(True)
        self.data_thread.start()

    def update_tree(self):
        b_prj_info = self.with_prj_info.get()
        b_prj_para = self.with_prj_info.get()

        l_wbs = len(self.wbs_bom)

        logger.info("正在更新WBS BOM列表...")
        for i in range(1, l_wbs + 1):
            self.wbs_list.insert('', END, values=self.wbs_bom[i])
        logger.info("WBS BOM列表更新完成.")

        if b_prj_info:
            logger.info("正在更新项目基本信息列表...")
            l_prj_info = len(self.prj_info_st)

            for i in range(1, l_prj_info + 1):
                self.prj_info_list.insert('', END, values=self.prj_info_st[i])
            logger.info("项目基本信息列表更新完成。")

        if b_prj_para:
            logger.info("正在更新项目参数信息列表...")
            l_prj_para = len(self.prj_para_st)

            for i in range(1, l_prj_para+1):
                self.prj_para_list.insert('', END, values=self.prj_para_st[i])
            logger.info("项目参数信息列表更新完成.")
