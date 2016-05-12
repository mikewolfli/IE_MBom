#!/usr/bin/env python3
#coding:utf-8
"""
  Author:   10256603<mikewolf.li@tkeap.com>
  Purpose: 
  Created: 2016/4/7
"""
from global_list import *
global login_info

header_line1 = ['TDM_ID','CN_DRAWING_NUMBER','CN_OLD_DRAWING_NUMBER','CN_PART_NAME_CHINESE','TDM_DESCRIPTION',
                'CN_OUTLINE_SIZE','CN_PART_TYPE|TN_PART_TYPE|TDM_NAME','CN_MATERIAL_LOC','CN_PART_WEIGHT','TDM_UOM|TDM_UNIT_OF_MEASURE|TDM_NAME',
                'CN_COMMENT','CN_COMPONENT_GROUP_CODE_SJ|TN_SYSTEMS|Cn_code_value','CN_COMPONENT_GROUP_CODE_ZS|TN_SYSTEMS|Cn_code_value',
                'CN_ACTUAL_BOX_SJ|TN_SYSTEMS|Cn_code_value','CN_ACTUAL_BOX_ZS|TN_SYSTEMS|Cn_code_value','CN_IN_EO_NO']
header_line2 = ['ID--Part','Drawing Number--Part','Old Drawing Number--Part','Part Name Chinese--Part','Description--Part','Outline Size--Part',
                'Item Type--Part','Material Loc--Part','Item Weight--Part','Units of Measure--Part','Comment--Part']

col_header = ['非标物料申请编号','判断','物料号','物料名称(中)','物料名称(英)','图号','单位','备注','RP','BoxId','申请人','旧物料号']
mat_db_header =['nstd_app','mat_no','mat_name_cn','mat_name_en','drawing_no','mat_unit','comments','rp','box_code_sj','justify','app_person','old_mat_no']
class import_pane(Frame):
    '''
    权限
    2- 导入物料
    3- 物料分类
    '''
    mr_rows = 0
    mat_dic= {} 
    
    cs_set = []
    nstd_dic ={}
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.widgets_group1 = Frame(self)
        self.widgets_group1.grid(row=0, column=0, columnspan=2, sticky=NSEW)
        self.im_button = Button(self.widgets_group1, text='导入数据')
        self.im_button.pack(side='left')
        self.im_button['command']=self.import_data
        self.im_cs_button = Button(self.widgets_group1, text='导入CS数据')
        self.im_cs_button.pack(side='left')
        self.im_cs_button['command']=self.import_data
        self.out_button = Button(self.widgets_group1, text='导出PDM导入表')
        self.out_button.pack(side='left')
        self.out_button['command']=self.out_pdm_excel
        self.del_button = Button(self.widgets_group1, text='删除物料(by申请号)')
        self.del_button.pack(side='left')
        self.del_button['command']=self.del_mats
        if login_info['perm'][0]!='2' and login_info['perm'][0]!='9':
            self.widgets_group1.grid_forget()  

        self.widgets_group2=Frame(self)
        self.widgets_group2.grid(row=0,column=2, columnspan=4, sticky=NSEW) 
        self.clu_p =Button(self.widgets_group2, text='CLU外购')
        self.clu_p.pack(side='left')
        self.clu_p['command']=self.clu_p_button
        self.pcu_p = Button(self.widgets_group2, text='PCU外购')
        self.pcu_p.pack(side='left')
        self.pcu_p['command']=self.pcu_p_button
        self.f30_p = Button(self.widgets_group2, text='F30外购')
        self.f30_p.pack(side='left')
        self.f30_p['command']=self.f30_p_button
        self.elec_m = Button(self.widgets_group2, text='电气自制')
        self.elec_m.pack(side='left')
        self.elec_m['command']=self.elec_m_button
        self.metal_m = Button(self.widgets_group2, text='钣金自制')
        self.metal_m.pack(side='left')
        self.metal_m['command']=self.metal_m_button
        self.tm_m = Button(self.widgets_group2, text='曳引机自制')
        self.tm_m.pack(side='left')
        self.tm_m['command']=self.tm_m_button
        if login_info['perm'][0]!='3' and login_info['perm'][0]!='9':
            self.widgets_group2.grid_forget()

        self.wl_str = StringVar()
        self.wl_entry = Entry(self, textvariable=self.wl_str)
        self.wl_entry.grid(row=0, column=6, sticky = EW)
        self.wl_entry.bind("<Return>", self.wl_result)
        self.mat_list = TableModel()

        for colname in col_header:
            self.mat_list.addColumn(colname)
        self.mat_list.addRow()
        table_panel = Frame(self)
        self.mat_table = TableCanvas(table_panel, self.mat_list, cellwidth=150, 
                                     cellbackgr='#e3f698',rowheaderwidth=30,
                                     rowselectedcolor='yellow', editable=False)
        self.mat_table.createTableFrame()

        self.mat_table.grid(row=0, column=0, rowspan=2, columnspan=2, sticky=NSEW)
        table_panel.columnconfigure(1, weight=1)
        table_panel.rowconfigure(1, weight=1)
        table_panel.grid(row=1, column=0, rowspan=2, columnspan=8, sticky=NSEW )
        self.grid(row=0, column=0, sticky=NSEW)
        self.columnconfigure(7, weight=1)
        self.rowconfigure(2, weight=1)         

    def import_data(self):
        file_list = filedialog.askopenfilenames(title="导入文件", filetypes=[('excel file','.xlsx'),('excel file','.xlsm')])
        if not file_list:
            return

        self.mr_rows = 0
        self.mat_dic={}
        self.mat_list.deleteRows()
        self.nstd_app_id = []
        for file in file_list:
            self.read_excel_file(file)
        self.mat_list.importDict(self.mat_dic) 
        self.mat_table.createTableFrame()

    def read_excel_file(self, file):
        wb = load_workbook(file, read_only=True, data_only=True)
        sheetnames=wb.get_sheet_names()
        
        if sheetnames[0].find('物料汇总表') !=-1:
            self.read_workbook(wb, sheetnames[0])
        else:
            self.read_workbook_cs(wb, sheetnames[0])
    
    def read_workbook_cs(self, wb, sheetname):
        ws1 = wb.get_sheet_by_name(sheetname)
        self.cs_set=[]
        self.nstd_dic={}
        rows = ws1.max_row-1
        for rx in range(rows):
            mr_line = {} 
            r_str = ws1.cell(row=rx+4, column=7).value
 
            if r_str is None:
                break
            
            s_str = str(r_str).strip()
            mr_line[col_header[0]] = ''
            mr_line[col_header[1]]=Justify_Types[0]
            mr_line[col_header[2]]=s_str
            r_str = ws1.cell(row=rx+4, column=10).value
            s_str = self.conver_excel_data(r_str)
            mr_line[col_header[3]] = s_str
            r_str = ws1.cell(row=rx+4, column=18).value
            s_str = self.conver_excel_data(r_str)
            mr_line[col_header[4]] = s_str
            r_str = ws1.cell(row=rx+4, column=14).value
            s_str = self.conver_excel_data(r_str)
            mr_line[col_header[5]] = s_str
            r_str = ws1.cell(row=rx+4, column=19).value
            s_str = self.conver_excel_data(r_str)
            mr_line[col_header[6]] = s_str
            r_str = ws1.cell(row=rx+4, column=12).value
            s_str = self.conver_excel_data(r_str)
            mr_line[col_header[7]] = s_str
            r_str = ws1.cell(row=rx+4, column=17).value
            s_str = self.conver_excel_data(r_str)
            mr_line[col_header[8]] = s_str
            r_str = ws1.cell(row=rx+4, column=15).value
            s_str = self.conver_excel_data(r_str)
            mr_line[col_header[9]] = self.boxid_convert(s_str)
            r_str = ws1.cell(row=rx+4, column=9).value
            s_str = self.conver_excel_data(r_str)            
            mr_line[col_header[10]]=s_str
            r_str = ws1.cell(row=rx+4, column=4).value
            s_str = self.conver_excel_data(r_str)
            mr_line['wbs']=s_str
            r_str = ws1.cell(row=rx+4, column=22).value
            s_str = self.conver_excel_data(r_str)  
            mr_line[col_header[11]]=s_str
            self.cs_set.append(mr_line)
            self.mat_dic[self.mr_rows+1]=mr_line
            self.mr_rows+=1
            
        self.scan_list()
        for item in self.cs_set:
            self.import_cs(item)      
        
    def import_cs(self, item):
        try: 
            nstd_mat_table.get(nstd_mat_table.mat_no==item[col_header[2]])
        except nstd_mat_table.DoesNotExist:
            try:
                nstd_app_head.get(nstd_app_head.nstd_app==item[col_header[0]])
            except nstd_app_head.DoesNotExist:
                nstd_id = item[col_header[0]]
                if len(item['wbs'])==0:
                    prj = ''
                else:
                    prj = item['wbs'][0:10]
                        
                nstd_app_head.create(nstd_app=nstd_id, project=prj, index_mat='NONCE') 
                
                if len(item['wbs'])!=0:
                    for wbs in self.nstd_dic[nstd_id]:
                        nstd_app_link.get_or_create(nstd_app=nstd_id, wbs_no=wbs, mbom_fin=False)                     
            
            nstd_mat_table.create(mat_no=item[col_header[2]], mat_name_cn=item[col_header[3]],
                                  mat_name_en=item[col_header[4]], drawing_no=item[col_header[5]],
                                  mat_unit=item[col_header[6]],comments=item[col_header[7]],
                                  rp=item[col_header[8]],box_code_sj=item[col_header[9]],
                                  nstd_app = item[col_header[0]], mat_app_person=item[col_header[10]], old_mat_no=item[col_header[11]])

            try:
                nstd_mat_fin.get(nstd_mat_fin.mat_no==item[col_header[2]])
            except nstd_mat_fin.DoesNotExist:
                nstd_mat_fin.create(mat_no=item[col_header[2]], justify=value2key(Justify_Types,item[col_header[1]]), mbom_fin=False,\
                                    pu_price_fin=False, co_run_fin=False, modify_by=login_info['uid'], modify_on=datetime.datetime.now())                     
        
    def scan_list(self):
        if len(self.cs_set)==0:
            return
        
        self.cs_set.sort(key= lambda x:(x['wbs']))
        
        project= None
        i_nstd = self.get_nstd_id_for_noce()
        wbses=[]
        i=0
        for item in self.cs_set:            
            if len(item['wbs'])==0:
                if project == item['wbs']:
                    item[col_header[0]]=nstd_id
                else:
                    if i!=0  and len(wbses)>0:
                        self.nstd_dic[nstd_id]=wbses
                        wbses=[]  
                    i_nstd+=1
                    nstd_id = 'NONCE'+str(i_nstd)+'-WL'
                    item[col_header[0]]=nstd_id
                    project = item['wbs'] 
                    self.nstd_dic[nstd_id] = None
            else:
                if project == item['wbs'][0:10]:
                    item[col_header[0]]=nstd_id
                    wbses.append(item['wbs']) 
                else:
                    if i!=0:
                        self.nstd_dic[nstd_id]=wbses
                        wbses=[]
                    i_nstd+=1
                    nstd_id = 'NONCE'+str(i_nstd)+'-WL'
                    item[col_header[0]]=nstd_id
                    project = item['wbs'][0:10]
                    wbses.append(item['wbs'])   
            i+=1
            
        if len(nstd_id)!=0 and len(wbses)!=0:
            self.nstd_dic[nstd_id]=wbses
            
    def boxid_convert(self, s):
        i_pos = s.find('#')
        if i_pos == -1 or s.isdecimal() or len(s)==0:
            return s
        
        e=''
        b=False
        e1=''
        for i in range(len(s)):
            if s[i].isdecimal():
                if len(e)<2:
                    e=e+s[i]
                else:
                    e1=e1+s[i]
            else:
                if len(e)>0 and len(e)<2:
                    e='0'+e               
                if len(e1)>0:
                    break
        
        if len(e)==1:
            e='0'+e
            
        for j in range(3-len(e1)):
            e1='0'+e1
            
        return e+e1
                    
    def get_nstd_id_for_noce(self):
        return  nstd_app_head.select().where(nstd_app_head.index_mat=='NONCE').count()    
               
    def read_workbook(self, wb, sheetname):
        ws1 = wb.get_sheet_by_name(sheetname)
        #ws2 =  wb.get_sheet_by_name(sheetnames[2])

        nstd_id1 = str(ws1['H6'].value).strip()
        nstd_id2 = str(ws1['I6'].value).strip()

        if nstd_id1.startswith('CE'):
            nstd_id = nstd_id1.rstrip()
            b_sel=0
        elif nstd_id2.startswith('CE'):
            nstd_id = nstd_id2.rstrip()
            b_sel=1
        else:
            return False

        if not nstd_id.endswith('-WL'):
            i_pos = nstd_id.find('WL')
            nstd_id=nstd_id[:i_pos+1]

        self.nstd_app_id.append(nstd_id)
        try:
            nstd_result = NonstdAppItem.select(NonstdAppItem.link_list, NonstdAppItemInstance.index_mat, NonstdAppItemInstance.res_engineer, NonstdAppItemInstance.create_emp).join(NonstdAppItemInstance, on=(NonstdAppItem.index == NonstdAppItemInstance.index))\
                .where((NonstdAppItemInstance.nstd_mat_app==nstd_id)&(NonstdAppItem.status>=0)&(NonstdAppItemInstance.status>=0)).naive().get()
        except NonstdAppItem.DoesNotExist:
            return False

        wbs_res = nstd_result.link_list
        index_mat = nstd_result.index_mat
        app_engineer_id= nstd_result.res_engineer

        if not app_engineer_id:
            app_engineer_id= nstd_result.create_emp

        try:
            emp_res= SEmployee.get(SEmployee.employee==app_engineer_id)
            app_per = emp_res.name
        except SEmployee.DoesNotExist:
            app_per=''           

        i_pos = index_mat.find('-')
        nstd_app_id = index_mat[0:i_pos]
        try:
            nstd_app_result=NonstdAppHeader.get((NonstdAppHeader.nonstd==nstd_app_id)&(NonstdAppHeader.status>=0))
        except NonstdAppHeader.DoesNotExist:
            return False

        project_id = nstd_app_result.project
        contract_id = nstd_app_result.contract

        try:
            nstd_app_head.get(nstd_app_head.nstd_app == nstd_id)
            q= nstd_app_head.update(project=project_id, contract=contract_id, index_mat=index_mat, app_person=app_per).where(nstd_app_head.nstd_app == nstd_id)
            q.execute()
        except nstd_app_head.DoesNotExist:
            nstd_app_head.create(nstd_app=nstd_id, project=project_id, contract=contract_id, index_mat=index_mat, app_person=app_per)

        if isinstance(wbs_res, str): 
            wbs_list = wbs_res.split(';')
        else:
            wbs_list=['']

        for wbs in wbs_list:  
            if  len(wbs.strip())==0 and len(wbs_list)>1:
                continue
            nstd_app_link.get_or_create(nstd_app=nstd_id, wbs_no=wbs[0:14], mbom_fin=False) 

        rows = ws1.max_row-1
        for rx in range(rows):
            mr_line = {}

            r_str = ws1.cell(row=rx+11, column=11+b_sel).value

            if r_str is None:
                break
            s_str = str(r_str).strip()
            mr_line[col_header[0]]=nstd_id
            mr_line[col_header[1]]=Justify_Types[0]
            mr_line[col_header[2]]=s_str
            r_str = ws1.cell(row=rx+11, column=9+b_sel).value
            s_str = self.conver_excel_data(r_str)
            mr_line[col_header[3]] = s_str
            r_str = ws1.cell(row=rx+11, column=10+b_sel).value
            s_str = self.conver_excel_data(r_str)
            mr_line[col_header[4]] = s_str
            r_str = ws1.cell(row=rx+11, column=13+b_sel).value
            s_str = self.conver_excel_data(r_str)
            mr_line[col_header[5]] = s_str
            r_str = ws1.cell(row=rx+11, column=26+b_sel).value
            s_str = self.conver_excel_data(r_str)
            mr_line[col_header[6]] = s_str
            r_str = ws1.cell(row=rx+11, column=16+b_sel).value
            s_str = self.conver_excel_data(r_str)
            mr_line[col_header[7]] = s_str
            r_str = ws1.cell(row=rx+11, column=23+b_sel).value
            s_str = self.conver_excel_data(r_str)
            mr_line[col_header[8]] = s_str
            r_str = ws1.cell(row=rx+11, column=21+b_sel).value
            s_str = self.conver_excel_data(r_str)
            mr_line[col_header[9]] = s_str
            mr_line[col_header[10]]=app_per
            mr_line[col_header[11]]=''

            try:
                nstd_mat_table.get(nstd_mat_table.mat_no==mr_line[col_header[2]])
            except nstd_mat_table.DoesNotExist:
                nstd_mat_table.create(mat_no=mr_line[col_header[2]], mat_name_cn=mr_line[col_header[3]],
                                      mat_name_en=mr_line[col_header[4]], drawing_no=mr_line[col_header[5]],
                                      mat_unit=mr_line[col_header[6]],comments=mr_line[col_header[7]],
                                      rp=mr_line[col_header[8]],box_code_sj=mr_line[col_header[9]],
                                      nstd_app = mr_line[col_header[0]], mat_app_person=mr_line[col_header[10]])

            try:
                nstd_mat_fin.get(nstd_mat_fin.mat_no==mr_line[col_header[2]])
            except nstd_mat_fin.DoesNotExist:
                nstd_mat_fin.create(mat_no=mr_line[col_header[2]], justify=value2key(Justify_Types,mr_line[col_header[1]]), mbom_fin=False,\
                                    pu_price_fin=False, co_run_fin=False, modify_by=login_info['uid'], modify_on=datetime.datetime.now())
            self.mat_dic[self.mr_rows+1]=mr_line

            self.mr_rows+=1

        return True

    def conver_excel_data(self, data):
        if (data is None) or (data=='N/A') or (data=='0') or (data=='N') or (data=='无'):
            return ''
        else:
            return str(data).strip()

    def del_mats(self):
        nstd_id = simpledialog.askstring('删除物料','请输入完整的物料申请号(不区分大小写)')
        if not nstd_id:
            return

        mat_items = nstd_mat_table.select().join(nstd_mat_fin).where((nstd_mat_table.nstd_app==nstd_id.upper())&(nstd_mat_fin.justify>0)).naive()

        if mat_items:
            messagebox.showwarning('警告','非标申请物料无法删除，因为已经提交分类')
            return

        del_qer = nstd_app_head.delete().where(nstd_app_head.nstd_app==nstd_id.upper())
        r = del_qer.execute()

        if i>0:
            messagebox.showinfo('提示','删除成功')

    def out_pdm_excel(self):
        file_str=filedialog.asksaveasfilename(title="导出文件", initialfile="temp",filetypes=[('excel file','.xlsx')])
        if not file_str:
            return

        if not file_str.endswith(".xlsx"):
            file_str+=".xlsx"

        wb=Workbook()
        ws=wb.worksheets[0]
        ws.title='物料导入表'
        for i in range(0,len(header_line1)):
            ws.cell(row=1,column=i+1).value=header_line1[i]  

        for i in range(0,len(header_line2)):
            ws.cell(row=2,column=i+1).value=header_line2[i]

        for i in range(1, self.mr_rows+1):
            ws.cell(row=i+2, column=1).value = self.mat_dic[i][col_header[2]]
            ws.cell(row=i+2, column=2).value = self.mat_dic[i][col_header[5]]
            ws.cell(row=i+2, column=4).value = self.mat_dic[i][col_header[3]]
            ws.cell(row=i+2, column=5).value = self.mat_dic[i][col_header[4]]
            ws.cell(row=i+2, column=10).value = self.mat_dic[i][col_header[6]]
            ws.cell(row=i+2, column=11).value = self.mat_dic[i][col_header[7]]
            ws.cell(row=i+2, column=12).value = self.mat_dic[i][col_header[8]]
            ws.cell(row=i+2, column=14).value = self.mat_dic[i][col_header[9]]
            ws.cell(row=i+2, column=3).value = self.mat_dic[i][col_header[11]]
            
        if excel_xlsx.save_workbook(workbook=wb, filename=file_str):
            messagebox.showinfo("输出","成功输出!")  

    def wl_result(self, event):
        self.mr_rows = 0
        self.mat_dic.clear()
        self.mat_list.deleteRows()
        s_str = self.wl_str.get()

        if s_str:
            s_result = nstd_app_head.select(nstd_app_head, nstd_mat_table, nstd_mat_fin).join(nstd_mat_table).switch(nstd_mat_table).join(nstd_mat_fin).where(nstd_app_head.nstd_app.contains(s_str)).order_by(nstd_app_head.nstd_app.asc(),nstd_mat_table.mat_no.asc()).naive()
            if s_result:
                for r in s_result: 
                    mr_line = {}  
                    mr_line[col_header[0]]= r.nstd_app
                    mr_line[col_header[1]]= Justify_Types[r.justify]
                    mr_line[col_header[2]]= r.mat_no
                    mr_line[col_header[3]]= r.mat_name_cn
                    mr_line[col_header[4]]= r.mat_name_en
                    mr_line[col_header[5]]= r.drawing_no
                    mr_line[col_header[6]]= r.mat_unit
                    mr_line[col_header[7]]= r.comments
                    mr_line[col_header[8]]= r.rp
                    mr_line[col_header[9]]= r.box_code_sj
                    
                    r_str = r.app_person
                    if isinstance(r_str, str):
                        mr_line[col_header[10]]= r_str
                    else:
                        mr_line[col_header[10]] = r.mat_app_person
                    
                    mr_line[col_header[11]]=r.old_mat_no
                        
                    self.mat_dic[self.mr_rows+1]=mr_line
                    self.mr_rows+=1

        self.mat_list.importDict(self.mat_dic) 
        self.mat_table.createTableFrame()

    def clu_p_button(self):
        #self.set_justify(1)
        d = ask_list('物料拷贝器')
        if not d:
            return

        self.set_justify(1, d)

    def pcu_p_button(self):
        #self.set_justify(2)
        d = ask_list('物料拷贝器')
        if not d:
            return

        self.set_justify(2, d)  
        
    def f30_p_button(self):
        d = ask_list('物料拷贝器')
        if not d:
            return

        self.set_justify(6, d)          

    def elec_m_button(self):
        #self.set_justify(3)
        d = ask_list('物料拷贝器')
        if not d:
            return

        self.set_justify(3, d)      

    def metal_m_button(self):
        #self.set_justify(4)
        d = ask_list('物料拷贝器')
        if not d:
            return

        self.set_justify(4, d)  

    def tm_m_button(self):
        #self.set_justify(5)
        d = ask_list('物料拷贝器')
        if not d:
            return

        self.set_justify(5, d)

    def set_justify(self, choice, mats):        
        self.mr_rows=0
        self.mat_dic={}
        self.mat_list.deleteRows()

        for mat in mats:
            try:
                q = nstd_mat_fin.get(nstd_mat_fin.mat_no==mat)
                
                if q.justify==0 or ((q.justify ==1 or q.justify==2 or q.justify==6) and not q.pu_price_fin) or \
                  ((q.justify==3 or q.justify==4 or q.justify==5) and not q.mbom_fin):                   
                    q=nstd_mat_fin.update(justify=choice).where(nstd_mat_fin.mat_no==mat)
                    i_r=q.execute() 

                    if i_r==0:
                        continue
            except nstd_mat_fin.DoesNotExist:
                continue

            r = nstd_app_head.select(nstd_app_head, nstd_mat_table, nstd_mat_fin).join(nstd_mat_table).switch(nstd_mat_table).join(nstd_mat_fin).where(nstd_mat_table.mat_no==mat).naive().get()
            if r:
                mr_line = {}
                mr_line[col_header[0]]= r.nstd_app
                mr_line[col_header[1]]= Justify_Types[r.justify]
                mr_line[col_header[2]]= r.mat_no
                mr_line[col_header[3]]= r.mat_name_cn
                mr_line[col_header[4]]= r.mat_name_en
                mr_line[col_header[5]]= r.drawing_no
                mr_line[col_header[6]]= r.mat_unit
                mr_line[col_header[7]]= r.comments
                mr_line[col_header[8]]= r.rp
                mr_line[col_header[9]]= r.box_code_sj
                if len(r.app_person)==0 or r.app_person=='\'\'':
                    mr_line[col_header[10]]=r.mat_app_person
                else:
                    mr_line[col_header[10]]=r.app_person
                self.mat_dic[self.mr_rows+1]=mr_line
            self.mr_rows+=1 
        self.mat_list.importDict(self.mat_dic) 
        self.mat_table.createTableFrame()                        

    ''' 
    def set_justify(self, val):
        file_list = filedialog.askopenfilenames(title="导入文件", filetypes=[('excel file','.xlsx'),('excel file','.xlsm')])
        if not file_list:
            return 

        for file in file_list:
            wb = load_workbook(file, read_only=True)
            sheetnames=wb.get_sheet_names()

        ws =  wb.get_sheet_by_name(sheetnames[1])
        rows = ws.max_row-1
        self.mr_rows = 0
        self.mat_dic.clear()        
        for rx in range(rows):
            r_str = ws.cell(row=rx+3,column=7).value

            if r_str is None:
                break

            s_str =str(r_str).rstrip()
            q=nstd_mat_fin.update(justify=val).where(nstd_mat_fin.mat_no==s_str)
            i_r=q.execute() 

            if i_r==0:
                continue

            r = nstd_mat_table.select(nstd_mat_table, nstd_mat_fin).join(nstd_mat_fin).where(nstd_mat_table.mat_no==s_str).get()
            if r:
                mr_line = {}
                mr_line[col_header[0]]= r.nstd_app
                mr_line[col_header[1]]= Justify_Types[r.nstd_mat_fin.justify]
                mr_line[col_header[2]]= r.mat_no
                mr_line[col_header[3]]= r.mat_name_cn
                mr_line[col_header[4]]= r.mat_name_en
                mr_line[col_header[5]]= r.drawing_no
                mr_line[col_header[6]]= r.mat_unit
                mr_line[col_header[7]]= r.comments
                mr_line[col_header[8]]= r.rp
                mr_line[col_header[9]]= r.box_code_sj
                self.mat_dic[self.mr_rows+1]=mr_line
                self.mr_rows+=1 
        self.mat_list.importDict(self.mat_dic) 
        self.mat_table.createTableFrame()  
    '''


threadLock = threading.Lock()
class refresh_thread(threading.Thread):
    def __init__(self, pane, typ=None):
        threading.Thread.__init__(self)
        self.pane=pane
        self.type=typ

    def run(self):
        threadLock.acquire()
        self.pane.refresh()
        threadLock.release()

display_col=['col0','col1','col2','col3','col4','col5','col6','col7','col8','col9','col10','col11','col21','col22','col23','col24','col25','col26']
cols = ['col0','col1','col2','col3','col4','col5','col6','col7','col8','col9','col10','col11','col12','col13','col14','col15','col16','col17','col18','col19','col20','col21','col22','col23','col24','col25','col26']
tree_head=['导入日期','非标编号','判断','物料号','物料名称(中)','物料名称(英)','图号','单位','备注','RP','BoxId','申请人','自制完成','负责人','完成日期','价格维护','负责人','完成日期','CO-run','负责人','完成日期','是否有图纸','图纸上传','IN SAP','是否紧急','项目交货期','要求配置完成日期']                          

class mat_fin_pane(Frame):
    '''
    所有符合条件的数据存储在mat_res这个字典中
    字典结构：
     mat_res-{mat_no：[第一级value1],...}
     nstd_res-{nstd: [wbs_no_list],...}
     wbs_res-{wbs_no: [unit_info],....}

    '''
    mat_res={}
    nstd_res={}
    wbs_res={} #存储unit wbs相关信息， key-wbs no, value-列表[wbs no, 合同号， 项目名称， 梯号，梯型，载重，速度，非标类型，是否紧急，项目交货期，要求配置完成日期]
    nstd_temp_res={}#存储部分计算值， key=nstd_app, value-列表[图纸上传，in SAP, 是否紧急， 项目交货期，要求配置完成日期]
    def __init__(self, master=None):
        Frame.__init__(self, master)  

        box_id_head = Label(self, text='箱号',anchor='w')
        box_id_head.grid(row=0,column=0, sticky=W)

        self.boxid_var = StringVar()
        self.boxid_cb = ttk.Combobox(self, textvariable= self.boxid_var, values=self.get_list_for(0), state='readonly')
        self.boxid_cb.grid(row=1,column=0, sticky=EW)
        self.boxid_var.set('All')
        self.boxid_cb.bind('<<ComboboxSelected>>', self.filter_list)

        rp_head = Label(self, text='RP', anchor='w')
        rp_head.grid(row=0, column=1, sticky=W)
        self.rp_var = StringVar()
        self.rp_cb = ttk.Combobox(self, textvariable=self.rp_var, values=self.get_list_for(1), state='readonly')
        self.rp_var.set('All')
        self.rp_cb.grid(row=1, column=1, sticky=EW)
        self.rp_cb.bind('<<ComboboxSelected>>', self.filter_list)

        mat_head = Label(self, text='物料号搜索', anchor='w')
        mat_head.grid(row=0, column=2, sticky=W)
        self.mat_var = StringVar()
        self.mat_entry=Entry(self, textvariable=self.mat_var)
        self.mat_entry.grid(row=1, column=2, sticky=EW)
        self.mat_entry.bind('<KeyRelease>', self.valid_mat)
        self.mat_entry.bind("<Return>", self.mat_search)

        self.elec_mt_button=Button(self, text='电气完成')
        self.elec_mt_button.grid(row=0, column=3, sticky=EW)
        self.elec_mt_button['command']=self.elec_fin
        if login_info['perm'][1]!='2' and login_info['perm'][1]!='9':
            self.elec_mt_button.grid_forget()

        self.metal_mt_button=Button(self, text='钣金完成')
        self.metal_mt_button.grid(row=0, column=4, sticky=EW)
        self.metal_mt_button['command']=self.metal_fin
        if login_info['perm'][1]!='3' and login_info['perm'][1]!='9':
            self.metal_mt_button.grid_forget()

        self.price_button=Button(self, text='价格维护完成')
        self.price_button.grid(row=1, column=3, sticky=EW)
        self.price_button['command']=self.price_fin
        if login_info['perm'][1]!='4' and login_info['perm'][1]!='9':
            self.price_button.grid_forget()    

        self.co_run_button=Button(self, text='CO-RUN完成')
        self.co_run_button.grid(row=1, column=4, sticky=EW)
        self.co_run_button['command']=self.co_run_fin
        if login_info['perm'][1]!='5' and login_info['perm'][1]!='9':
            self.co_run_button.grid_forget()

        self.tm_mt_button=Button(self, text='曳引机完成')
        self.tm_mt_button.grid(row=0, column=5, sticky=EW)
        self.tm_mt_button['command']=self.traction_fin
        if login_info['perm'][1]!='6' and login_info['perm'][1]!='9':
            self.tm_mt_button.grid_forget()

        self.refresh_button=Button(self, text='手动刷新')
        self.refresh_button.grid(row=1, column=6, sticky=EW)
        self.refresh_button['command']=self.__refresh_tree

        self.export_button=Button(self, text='导出EXCEL')
        self.export_button.grid(row=0, column=6,  sticky=EW)
        self.export_button['command']=self.__export_excel

        if login_info['perm'][1]=='1' or login_info['perm'][1]=='5' or login_info['perm'][1]=='9':
            display_col.insert(-6,'col12')
            display_col.insert(-6,'col13')
            display_col.insert(-6,'col14')
        if login_info['perm'][1]=='1' or login_info['perm'][1]=='5' or login_info['perm'][1]=='9':
            display_col.insert(-6,'col15')
            display_col.insert(-6,'col16')
            display_col.insert(-6,'col17')            
        if login_info['perm'][1]=='1' or login_info['perm'][1]=='9':
            display_col.insert(-6,'col18')
            display_col.insert(-6,'col19')
            display_col.insert(-6,'col20')             

        self.mat_list = ttk.Treeview(self, columns=cols,displaycolumns=display_col,
                                     selectmode='extended')
        style = ttk.Style()
        style.configure("Treeview", font=('TkDefaultFont', 10))
        style.configure("Treeview.Heading", font=('TkDefaultFont', 9)) 
        self.mat_list.heading("#0",text='')       
        for col in cols:
            i = cols.index(col)
            #self.mat_list.heading(col, text=tree_head[i])
            self.mat_list.heading(col, text=tree_head[i] , command=lambda _col=col: treeview_sort_column(self.mat_list, _col, False))
        self.mat_list.column('#0', width=20)
        self.mat_list.column('col0', width=80, anchor='w')
        self.mat_list.column('col1', width=100, anchor='w')
        self.mat_list.column('col2', width=80, anchor='w')
        self.mat_list.column('col3', width=80, anchor='w')
        self.mat_list.column('col4', width=100, anchor='w')
        self.mat_list.column('col5', width=100, anchor='w')
        self.mat_list.column('col6', width=100, anchor='w')
        self.mat_list.column('col7', width=40, anchor='w')
        self.mat_list.column('col8', width=200, anchor='w')
        self.mat_list.column('col9', width=50, anchor='w')
        self.mat_list.column('col10', width=50, anchor='w')
        self.mat_list.column('col11', width=60, anchor='w')
        self.mat_list.column('col12', width=60, anchor='w')
        self.mat_list.column('col13', width=50, anchor='w')
        self.mat_list.column('col14', width=100, anchor='w')
        self.mat_list.column('col15', width=60, anchor='w')
        self.mat_list.column('col16', width=50, anchor='w')
        self.mat_list.column('col17', width=100, anchor='w')
        self.mat_list.column('col18', width=60, anchor='w')
        self.mat_list.column('col19', width=50, anchor='w')
        self.mat_list.column('col20', width=100, anchor='w')
        self.mat_list.column('col21', width=50, anchor='w')
        self.mat_list.column('col22', width=150, anchor='w')
        self.mat_list.column('col23', width=150, anchor='w')
        self.mat_list.column('col24', width=50, anchor='w')
        self.mat_list.column('col25', width=80, anchor='w')
        self.mat_list.column('col26', width=80, anchor='w')
        ysb = ttk.Scrollbar(self, orient='vertical', command=self.mat_list.yview)
        xsb = ttk.Scrollbar(self, orient='horizontal', command=self.mat_list.xview)  
        self.mat_list.grid(row=2, column=0, rowspan=2, columnspan=8, sticky='nsew')
        
        ysb.grid(row=2, column=8,rowspan=2, sticky='ns')
        xsb.grid(row=4, column=0, columnspan=8, sticky='ew')    
        self.grid()
        self.columnconfigure(7, weight=1)
        self.rowconfigure(3, weight=1)

         
        self.mat_list.bind('<Alt-c>', self.copy_mat_list)
        self.mat_list.bind('<Alt-C>', self.copy_mat_list)
        self.mat_list.bind('<Control-c>', self.copy_list)
        self.mat_list.bind('<Control-C>', self.copy_list) 
        
        self.focus_force()
        self.bind('<Control-a>', self.select_all)
        self.bind('<Control-A>', self.select_all)
        self.bind('<Escape>', self.clear_select)
        self.mat_list.bind('<Control-a>', self.select_all)
        self.mat_list.bind('<Control-A>', self.select_all)

        if login_info['status'] and int(login_info['perm'][1])>=1:
            self.__loop_refresh()
            
    def combine_wbs(self, li):
        li.sort()
        if len(li)>1:
            head = li[0]
        elif li is None:
            return ''
        elif len(li)==0:
            return ''
        else:
            return li[0]
        
        start = int(li[0][11:])
        j=1
        end = ''
        for i in range(1, len(li)):
            if int(li[i][11:]) == start+j:
                j+=1   
            else:
                if j>1:
                    head=head+'~'+end
                elif len(end)>0:
                    head = head+','+end
                
                if j>1:
                    head=head+','+li[i][11:]
                start=int(li[i][11:])

                j=1 
            end = li[i][11:]
            
        if j>1:
            head=head+'~'+end 
        else:
            head = head+','+end
        
        return head
                      
    def __export_excel(self):
        items = self.mat_list.get_children('')
        if not items:
            return

        file_str=filedialog.asksaveasfilename(title="导出文件", filetypes=[('excel file','.xlsx')])
        if not file_str:
            return

        if not file_str.endswith(".xlsx"):
            file_str+=".xlsx"

        wb=Workbook()
        ws=wb.worksheets[0]
        ws.title='物料清单'
        col_size = len(cols)
        for i in range(col_size):
            ws.cell(row=1,column=i+1).value=tree_head[i]  
        ws.cell(row=1, column=col_size+1).value = '关联WBS'
        n=0 
        
        dic_str = {}
        for item in items:
            if self.mat_list.parent(item) !='':
                continue
            
            for i in range(col_size):
                ws.cell(row=n+2, column=i+1).value = self.mat_list.item(item, 'values')[i]
            
            j=cols.index('col1')
            nstd_id = self.mat_list.item(item, 'values')[j]
            
            if nstd_id not in dic_str:
                s_str = self.combine_wbs(self.nstd_res[nstd_id])
                dic_str[nstd_id]=s_str
                
            ws.cell(row=n+2, column=col_size+1).value = dic_str[nstd_id]            
            n+=1

        if excel_xlsx.save_workbook(workbook=wb, filename=file_str):
            messagebox.showinfo("输出","成功输出!")  

    def metal_fin(self):
        #self.choice=1
        self.__mat_update(1)

    def elec_fin(self):
        #self.choice=1
        self.__mat_update(1)

    def price_fin(self):
        #self.choice=2
        self.__mat_update(2)

    def co_run_fin(self):
        #self.choice=3
        self.__mat_update(3)

    def traction_fin(self):
        #self.choice=1
        self.__mat_update(1)

    def filter_list(self, event):
        rp = self.rp_var.get()
        boxid= self.boxid_var.get()
        self.update_tree_data(col9=rp, col10=boxid)

    def refresh(self):
        self.rp_var.set('All')
        self.boxid_var.set('All')
        self.refresh_mat()   
        self.update_tree_data()

    def get_list_for(self, fu):
        if fu==0:
            fd = fn.Distinct(nstd_mat_table.box_code_sj)
        if fu==1:
            fd = fn.Distinct(nstd_mat_table.rp)

        res = nstd_mat_table.select(fd)       
        all_tub=['All']

        for r in res:
            if fu==0:
                all_tub.append(r.box_code_sj)
            if fu==1:
                all_tub.append(r.rp)

        return all_tub

    def get_name(self, pid):
        if pid=='' or not pid:
            return ''

        try: 
            r_name = s_employee.get(s_employee.employee==pid)
            s_name = r_name.name
        except s_employee.DoesNotExist:
            return 'None'

        return s_name

    def refresh_mat(self):
        '''
        刷新物料字典
        '''
        for row in self.mat_list.get_children():
            self.mat_list.delete(row) 
            
        self.mat_res={}
        self.wbs_res={}
        self.nstd_res={}
        self.nstd_temp_res={}

        if login_info['perm'][1]=='2':
            mat_items=nstd_app_head.select(nstd_app_head, nstd_mat_table, nstd_mat_fin).join(nstd_mat_table).switch(nstd_mat_table).join(nstd_mat_fin).where((nstd_mat_fin.justify==3)&(nstd_mat_fin.mbom_fin==False)).naive()
        elif login_info['perm'][1]=='3':
            mat_items=nstd_app_head.select(nstd_app_head, nstd_mat_table, nstd_mat_fin).join(nstd_mat_table).switch(nstd_mat_table).join(nstd_mat_fin).where((nstd_mat_fin.justify==4)&(nstd_mat_fin.mbom_fin==False)).naive()
        elif login_info['perm'][1]=='4':
            mat_items=nstd_app_head.select(nstd_app_head, nstd_mat_table, nstd_mat_fin).join(nstd_mat_table).switch(nstd_mat_table).join(nstd_mat_fin).where(((nstd_mat_fin.justify==1)|(nstd_mat_fin.justify==2)|(nstd_mat_fin.justify==6))&(nstd_mat_fin.pu_price_fin==False)).naive()
        elif login_info['perm'][1]=='5':
            mat_items=nstd_app_head.select(nstd_app_head, nstd_mat_table, nstd_mat_fin).join(nstd_mat_table).switch(nstd_mat_table).join(nstd_mat_fin)\
                .where(((((nstd_mat_fin.justify==1)|(nstd_mat_fin.justify==2)|(nstd_mat_fin.justify==6))&(nstd_mat_fin.pu_price_fin==True))|(((nstd_mat_fin.justify==3)|(nstd_mat_fin.justify==4)|(nstd_mat_fin.justify==5))&(nstd_mat_fin.mbom_fin==True)))&(nstd_mat_fin.co_run_fin==False)).naive()
        elif login_info['perm'][1]=='6':
            mat_items=nstd_app_head.select(nstd_app_head, nstd_mat_table, nstd_mat_fin).join(nstd_mat_table).switch(nstd_mat_table).join(nstd_mat_fin).where((nstd_mat_fin.justify==5)&(nstd_mat_fin.mbom_fin==False)).naive()  
        elif login_info['perm'][1]=='1' or login_info['perm'][1]=='9':
            mat_items=nstd_app_head.select(nstd_app_head, nstd_mat_table, nstd_mat_fin).join(nstd_mat_table).switch(nstd_mat_table).join(nstd_mat_fin).where((nstd_mat_fin.justify>=0)&(nstd_mat_fin.co_run_fin==False)).naive()            
        else: 
            return False

        if not mat_items:
            return False

        self.build_data_model(mat_items)

        return True

    def build_data_model(self, query_res):
        i=0
        for r in query_res:
            item=[]
            nstd_app_id = r.nstd_app
            index_mat_id = r.index_mat
            item.append(none2str(date2str(r.modify_on)))
            item.append(nstd_app_id)
            item.append(Justify_Types[r.justify])
            mat_id = r.mat_no
            item.append(mat_id)
            item.append(r.mat_name_cn)
            item.append(r.mat_name_en)
            item.append(none2str(r.drawing_no))
            item.append(r.mat_unit)
            item.append(none2str(r.comments))
            item.append(none2str(r.rp))
            item.append(none2str(r.box_code_sj))
            app_per = r.app_person
            if app_per is None:
                app_per = r.mat_app_person
            elif len(app_per)==0:
                app_per = r.mat_app_person
            
            item.append(app_per)
            temp = (r.mbom_fin and 'Y' or '')
            item.append(temp)
            if temp.upper()=='Y':
                item.append(get_name(r.mbom_fin_by))
                item.append(none2str(r.mbom_fin_on))
            else:
                item.append('')
                item.append('')
            temp=(r.pu_price_fin and 'Y' or '')
            item.append(temp)
            if temp.upper()=='Y':
                item.append(get_name(r.pu_price_fin_by))
                item.append(none2str(r.pu_price_fin_on))
            else:
                item.append('')
                item.append('')                
            temp = (r.co_run_fin and 'Y' or '')
            item.append(temp)
            if temp.upper()=='Y':
                item.append(get_name(r.co_run_fin_by))
                item.append(none2str(r.co_run_fin_on))
            else:
                item.append('')
                item.append('') 

            if not r.drawing_no or r.drawing_no=='':
                has_drawing = ''
            else:
                has_drawing = 'Y'
            item.append(has_drawing)

            wbs_list=self.get_wbs_list(nstd_app_id)

            if nstd_app_id not in self.nstd_temp_res:
                self.nstd_temp_res[nstd_app_id]=self.get_nstd_info(index_mat_id, wbs_list)

            res= self.nstd_temp_res[nstd_app_id]
            for j in range(len(res)):
                item.append(none2str(res[j]))

            self.mat_res[i]=item
            i+=1

    def get_nstd_info(self, index, wbs_list):
        if index=='' or not index:
            return ''

        item=[]
        try:
            result = LProcAct.select(LProcAct.finish_date).where((LProcAct.instance==index)&(LProcAct.action=='AT00000020')).get()
            s_date = result.finish_date
        except LProcAct.DoesNotExist:
            s_date=''

        item.append(s_date)

        try:
            result = LProcAct.select(LProcAct.finish_date).where((LProcAct.instance==index)&(LProcAct.action=='AT00000015')).get()
            s_date = result.finish_date
        except LProcAct.DoesNotExist:
            s_date=''

        item.append(s_date)
        req_conf_fin=None
        req_delivery=None
        urgent_flag=''

        for wbs_line in wbs_list:
            if len(wbs_line)>=20:
                if wbs_line[23]=='Y':
                    urgent_flag='Y'

                if  wbs_line[25] and req_conf_fin is None:
                    req_conf_fin=wbs_line[25]
                elif wbs_line[25] and req_conf_fin:
                    if wbs_line[25]< req_conf_fin:
                        req_conf_fin=wbs_line[25]

                if wbs_line[24] and req_delivery is None:
                    req_delivery=wbs_line[24]
                elif wbs_line[24] and req_delivery:
                    if wbs_line[24]< req_delivery:
                        req_delivery=wbs_line[24] 

        item.append(urgent_flag)
        item.append(none2str(req_delivery))
        item.append(none2str(req_conf_fin))                  

        return item

    def get_wbs_list(self, nstd):
        """

        :type nstd: object
        """
        result = nstd_app_head.select(nstd_app_head, nstd_app_link.wbs_no).join(nstd_app_link).where(nstd_app_head.nstd_app==nstd).order_by(nstd_app_link.wbs_no).naive()
        wbs_list = []
        wbs_res_list=[]
        for r in result:
            wbs_no = r.wbs_no
            if not isinstance(wbs_no, str) or (isinstance(wbs_no, str) and wbs_no==''):
                wbs_no = r.project
            if wbs_no not in self.wbs_res:
                temp = self.get_wbs_info(wbs_no)
                if temp:
                    self.wbs_res[wbs_no]=temp
            wbs_list.append(wbs_no)

            wbs_res_list.append(self.wbs_res[wbs_no])

        if nstd not in self.nstd_res:
            self.nstd_res[nstd]=wbs_list

        return wbs_res_list

    def get_wbs_info(self, wbs):
        wbs_info=[]
        if not wbs.startswith('E'):
            try:
                result=ProjectInfo.get(ProjectInfo.project==wbs)
                wbs_info.append(wbs)
                wbs_info.append('')
                wbs_info.append(result.project_name)
            except ProjectInfo.DoesNotExist:
                return None
            return wbs_info

        try:
            result = ProjectInfo.select(ProjectInfo.project, ProjectInfo.contract, ProjectInfo.project_name, UnitInfo.lift_no,UnitInfo.project_catalog, UnitInfo.nonstd_level,\
                                        UnitInfo.is_urgent,UnitInfo.req_configure_finish, UnitInfo.req_delivery_date,ElevatorTypeDefine.elevator_type, SUnitParameter.load,SUnitParameter.speed).join(UnitInfo, on=(UnitInfo.project==ProjectInfo.project))\
                .switch(UnitInfo).join(ElevatorTypeDefine).join(SUnitParameter, on=(SUnitParameter.wbs_no==UnitInfo.wbs_no)).where(UnitInfo.wbs_no==wbs).naive().get()
        except ProjectInfo.DoesNotExist:
            return wbs_info

        wbs_info.append(wbs)#0
        wbs_info.append(none2str(result.contract))#1
        wbs_info.append(none2str(result.project_name))#2
        wbs_info.append(none2str(result.lift_no))#3
        wbs_info.append(none2str(result.elevator_type))#4
        wbs_info.append(none2str(result.load))#5
        wbs_info.append(none2str(result.speed))#6
        wbs_info.append(Catalog_Types[result.project_catalog])#7
        wbs_info.append(Nonstd_Level[result.nonstd_level])#8
        for i in range(0, 14):
            wbs_info.append('')
        wbs_info.append((result.is_urgent and 'Y' or ''))
        if not result.req_delivery_date:
            wbs_info.append(result.req_delivery_date)
        else:
            wbs_info.append(result.req_delivery_date.date())
        if not result.req_configure_finish:
            wbs_info.append(result.req_configure_finish)
        else:
            wbs_info.append(result.req_configure_finish.date())

        return wbs_info

    def update_tree_data(self, **cases):
        for key in range(len(self.mat_res)):
            b_show=True 
            for case in cases:
                try:
                    i = cols.index(case)
                except ValueError:
                    continue

                if self.mat_res[key][i].find(cases[case])==-1 and self.mat_res[key][i]!=cases[case] and cases[case].upper()!='ALL':
                    b_show=False
                    break

            if not b_show: 
                continue

            parent = self.mat_list.insert('', END, values=self.mat_res[key]) 
            i = cols.index('col1')
            wbs_list = self.nstd_res[self.mat_res[key][i]]

            for wbs in  wbs_list:
                self.mat_list.insert(parent, END, values=self.wbs_res[wbs])

    def __refresh_tree(self):        
        data_thread= refresh_thread(self)
        data_thread.setDaemon(True)
        data_thread.start()        

    def __loop_refresh(self):
        self.__refresh_tree()
        self.mat_list.after(1800000, self.__loop_refresh)
        #self.mat_list.after(10000, self.__loop_refresh)

    def mat_search(self, event):
        s_str = self.mat_var.get()
        self.update_tree_data(col3=s_str)

    def valid_mat(self, a):
        l = list(self.mat_entry.get())
        for i in range(len(l) - 1, -1, -1):
            if not(48 <= ord(l[i]) <= 57):
                self.mat_entry.delete(i, i+1)

    def clear_select(self, event):
        items=self.mat_list.selection()
        for item in items:
            self.mat_list.selection_remove(item)

    def copy_mat_list(self, event):
        items=self.mat_list.selection()
        if not items:
            return
        mat_str=''
        self.mat_list.clipboard_clear()
        for item in items:
            if self.mat_list.parent(item)!='':
                continue
            i = cols.index('col3')
            mat_str= self.mat_list.item(item, 'values')[i]+'\n'
            self.mat_list.clipboard_append(mat_str)

    def copy_list(self, event):
        items=self.mat_list.selection()
        if not items:
            return 

        self.mat_list.clipboard_clear()
        mat_str=''

        for col in display_col:
            i=cols.index(col)
            mat_str = mat_str+tree_head[i]+'\t'
        mat_str = mat_str+'\n'
        self.mat_list.clipboard_append(mat_str)

        for item in items:
            if self.mat_list.parent(item)!='':
                continue        
            mat_str =''
            for col in display_col:
                j=cols.index(col)
                mat_str= mat_str + self.mat_list.item(item, 'values')[j]+'\t'
            mat_str=mat_str+'\n'
            self.mat_list.clipboard_append(mat_str)

    def select_all(self, event):      
        items = self.mat_list.get_children()
        if not items:
            return

        self.mat_list.selection_set(items[0])   
        self.mat_list.focus_set()
        self.mat_list.focus(items[0])
        #self.mat_list.selection_remove(items[0])
        for item in items:
            self.mat_list.selection_add(item)

    def copy_clip_process(self):
        wbs_sel = WBSSelDialog(self, 'WBS List粘贴')

    def __mat_list(self,choice, mats):
        s_error=''
        i_error=0
        i_suss=0
        for mat in mats:
            if self.process_mat_status(mat.rstrip(), choice, True)>0:
                i_suss=i_suss+1
            else:
                i_error=i_error+1
                s_error=s_error+mat.rstrip()+';'

        messagebox.showinfo('结果',str(i_suss)+'更新成功;'+str(i_error)+'更新失败:'+s_error)

        if i_suss>=1:
            self.__refresh_tree()

    def __mat_update(self, choice):
        items = self.mat_list.selection()
        if not items: 
            d = ask_list('物料拷贝器')
            if not d:
                return
            self.__mat_list(choice, d)            
            return

        if messagebox.askyesno('确认执行','此操作不可逆，确认执行(YES/NO)')==NO:
            return

        i_suss=0
        i_error=0
        s_error=''
        for item in items:
            if self.mat_list.parent(item) != '':
                continue
            col_index=cols.index('col3')
            mat = self.mat_list.item(item, 'values')[col_index]
            if self.process_mat_status(mat, choice, True)>0:
                i_suss=i_suss+1
            else:
                i_error=i_error+1
                s_error=s_error+mat+';'

        messagebox.showinfo('结果',str(i_suss)+'更新成功;'+str(i_error)+'更新失败:'+s_error)

        if i_suss>=1:
            self.__refresh_tree()

    def process_mat_status(self, mat, method, value=True):
        """
          method - 1 m_bom_fin 操作
                   2 pu_price_fin 操作
                   3 co_run 操作
          返回值 -  1 update成功
                   0 update不成功
                   -1 物料号不存在
                   -2 物料号分类和操作人员权限不符合
                   -3 co run的前一步未完成
                   2  物料条件已经是设定的值了
        """
        try:
            query_res= nstd_mat_fin.get(nstd_mat_fin.mat_no==mat)
        except nstd_mat_table.DoesNotExist:
            return -1

        mat_catalog= query_res.justify
        mat_mbom_fin= query_res.mbom_fin
        mat_pu_price_fin = query_res.pu_price_fin
        mat_co_run_fin = query_res.co_run_fin

        user_per = int(login_info['perm'][1])

        if method==1:
            #分类和权限组合判断
            if (mat_catalog==3 and (user_per==2 or user_per==9)) or \
               (mat_catalog==4 and (user_per==3 or user_per==9)) or \
               (mat_catalog==5 and (user_per==6 or user_per==9)):
                if mat_mbom_fin==value:
                    return 2
                else:
                    s = nstd_mat_fin.update(mbom_fin=value,mbom_fin_on=datetime.datetime.now(), mbom_fin_by=login_info['uid']).where(nstd_mat_fin.mat_no==mat)
                    i_u=s.execute()               
            else:
                return -2
        elif method==2:
            if (mat_catalog==1 or mat_catalog==2) and (user_per==4 or user_per==9):
                if mat_pu_price_fin ==value:
                    return 2
                else:
                    s = nstd_mat_fin.update(pu_price_fin=value,pu_price_fin_on=datetime.datetime.now(), pu_price_fin_by=login_info['uid']).where(nstd_mat_fin.mat_no==mat)
                    i_u=s.execute()  
            else: 
                return -2
        elif method==3:
            if user_per!=5 and user_per!=9:
                return -2

            if (mat_catalog==1 or mat_catalog==2) and not mat_pu_price_fin:
                return -3

            if (mat_catalog==3 or mat_catalog==4 or mat_catalog==5) and not mat_mbom_fin:
                return -3

            if mat_co_run_fin == value:
                return 2
            else:
                s = nstd_mat_fin.update(co_run_fin=value,co_run_fin_on=datetime.datetime.now(), co_run_fin_by=login_info['uid']).where(nstd_mat_fin.mat_no==mat)
                i_u=s.execute()              
        else:
            return method

        return i_u

project_cols=['col1','col2','col3','col4','col5','col6','col7','col8','col9','col10','col11','col12','col13','col14','col15','col16','col17']        
project_head = ['合同号','Unit WBS NO.','项目名称','梯号','梯型编号','梯型','Unit状态','是否紧急','项目类别','发货期','配置要求完成日期','分箱完成时间',
                '是否非标','非标预计完成日期','非标实际完成日期','是否有非标物料', '非标物料维护完成时间']
class proj_release_pane(Frame):
    '''    
    '''
    pos = None
    wbs_dic={} # 存储相关WBS NO相关信息， key- wbs no, value=[project_cols顺序存储的值]
    wbs_keys=[]
    func_list = ['配置完成未release项目','已经release项目','正在配置项目','重排产项目','未启动配置项目']
    def __init__(self, master=None):
        Frame.__init__(self, master)
        filter_body = Frame(self)
        filter_body.grid(row=0, column=0, rowspan=2, columnspan=2)
        self.sel_label=Label(filter_body, text='功能选择' )
        self.sel_label.grid(row=0, column=0, sticky=W)
        self.project_func = StringVar()
        self.func_sel = ttk.Combobox(filter_body, textvariable=self.project_func, values=self.func_list, state='readonly')
        self.func_sel.grid(row=1, column=0, sticky=W)
        self.func_sel.bind('<<ComboboxSelected>>', self.select_func)
        self.project_func.set(self.func_list[0])
        
        self.pos_label = Label(filter_body, text='项目定位')
        self.pos_label.grid(row=0, column=1, sticky=W)
        self.wbs_pos = StringVar()
        self.wbs_pos_input = Entry(filter_body, textvariable=self.wbs_pos )
        self.wbs_pos_input.grid(row=1, column=1, sticky=W)
        self.wbs_pos_input.bind("<Return>", self.wbs_search)
        
        self.refresh_button=Button(self, text='刷新数据')
        self.refresh_button.grid(row=1, column=2, sticky=NSEW)
        self.refresh_button['command']=self.__refresh_list
        
        self.release_button=Button(self, text='项目release')
        self.release_button.grid(row=0, column=2, sticky=NSEW)
        self.release_button['command']= self.release_prj
        if login_info['perm'][0]!='2' and login_info['perm'][0]!='9':
            self.release_button.grid_forget()  
        
        self.delivery_button=Button(self, text='发运完成')
        self.delivery_button.grid(row=1, column=3, sticky=NSEW)
        if login_info['perm'][0]!='3' and login_info['perm'][0]!='9':
            self.delivery_button.grid_forget()  
            
        info_label = Label(filter_body, text='红色背景-暂停项目;绿色背景-加急项目;\n黄色背景-精益项目;天蓝色:pre-engineer项目')
        info_label.grid(row=0, column=4, columnspan=3, sticky=NSEW)
        self.count_str = StringVar()
        counter_label = Label(filter_body, textvariable=self.count_str)
        counter_label.grid(row=1, column=4, sticky=W)        
        
        self.wbs_list = ttk.Treeview(self, show='headings', columns=project_cols,selectmode='extended')
        style = ttk.Style()
        style.configure("Treeview", font=('TkDefaultFont', 10))
        style.configure("Treeview.Heading", font=('TkDefaultFont', 9))       
        for col in project_cols:
            i = project_cols.index(col)
            self.wbs_list.heading(col, text=project_head[i] , command=lambda _col=col: treeview_sort_column(self.wbs_list, _col, False))
        
        self.wbs_list.column('col1', width=100, anchor='w')
        self.wbs_list.column('col2', width=150, anchor='w')
        self.wbs_list.column('col3', width=200, anchor='w')
        self.wbs_list.column('col4', width=60, anchor='w')
        self.wbs_list.column('col5', width=100, anchor='w')
        self.wbs_list.column('col6', width=80, anchor='w')
        self.wbs_list.column('col7', width=80, anchor='w')
        self.wbs_list.column('col8', width=80, anchor='w')
        self.wbs_list.column('col9', width=100, anchor='w')
        self.wbs_list.column('col10', width=100, anchor='w')
        self.wbs_list.column('col11', width=100, anchor='w')
        self.wbs_list.column('col12', width=150, anchor='w')
        self.wbs_list.column('col13', width=80, anchor='w')
        self.wbs_list.column('col14', width=150, anchor='w')
        self.wbs_list.column('col15', width=150, anchor='w')
        self.wbs_list.column('col16', width=150, anchor='w')
        self.wbs_list.column('col17', width=150, anchor='w')
        ysb = ttk.Scrollbar(self, orient='vertical', command=self.wbs_list.yview)
        xsb = ttk.Scrollbar(self, orient='horizontal', command=self.wbs_list.xview)  
        self.wbs_list.grid(row=2, column=0, rowspan=2, columnspan=9, sticky='nsew') 
        
        ysb.grid(row=2, column=9, rowspan=2, sticky='ns')
        xsb.grid(row=4, column=0, columnspan=9, sticky='ew')    
        self.grid()
        self.columnconfigure(8, weight=1)
        self.rowconfigure(3, weight=1)
        
        self.popup=Menu(self, tearoff=0)
        self.popup.add_command(label="显示配置工作流", command=self.com_configure)
        self.popup.add_command(label="显示非标申请工作流", command=self.com_nstd_app)
        self.popup.add_command(label="显示非标物料工作流", command=self.com_mat_maintain)
        self.popup.add_command(label="显示设计工作流", command=self.com_design)
        self.popup.add_separator()
        self.popup.add_command(label="显示非标选项", command=self.com_nonstd)
        self.popup.add_command(label="显示非标物料表", command=self.com_mats)
        
        self.wbs_list.bind("<Button-3>", self.do_popup)
        
        if login_info['status'] and int(login_info['perm'][2])>=1:
            self.__loop_refresh()
    
    def wbs_search(self, event):
        wbs = self.wbs_pos.get()
        if len(wbs)==0:
            return
        
        items = self.wbs_list.get_children()
        
        if not items:
            self.pos=None
            return
            
        for item in items:
            if self.pos and self.pos!=item:
                continue
            
            if self.pos==item:
                self.wbs_list.selection_remove(item)
                continue
        
            if self.wbs_list.item(item, 'values')[0].find(wbs)>=0 or self.wbs_list.item(item, 'values')[1].find(wbs)>=0:
                self.pos=item
                self.wbs_list.selection_set(item)
                return
        
        self.pos = None
        
    def com_configure(self):
        items = self.wbs_list.selection()
        
        if not items:
            return
        
        wbses = []
        for item in items:
            wbs = self.wbs_list.item(item, 'values')[1]
            wbses.append(wbs)
            
        his_display(self,'配置工作流', wbses, 0)

    def com_nstd_app(self):
        items = self.wbs_list.selection()
        
        if not items:
            return
        
        wbses = []
        for item in items:
            wbs = self.wbs_list.item(item, 'values')[1]
            wbses.append(wbs)
            
        his_display(self,'非标申请工作流', wbses, 1)
    
    def com_mat_maintain(self):
        items = self.wbs_list.selection()
        
        if not items:
            return
        
        wbses = []
        for item in items:
            wbs = self.wbs_list.item(item, 'values')[1]
            wbses.append(wbs)
            
        his_display(self,'非标物料维护工作流', wbses, 2)

    def com_design(self):
        items = self.wbs_list.selection()
        
        if not items:
            return
        
        wbses = []
        for item in items:
            wbs = self.wbs_list.item(item, 'values')[1]
            wbses.append(wbs)
            
        his_display(self,'非标设计工作流', wbses, 3)

    def com_nonstd(self):
        items = self.wbs_list.selection()
        
        if not items:
            return
        
        wbses = []
        for item in items:
            wbs = self.wbs_list.item(item, 'values')[1]
            wbses.append(wbs)
            
        his_display(self,'非标选项', wbses, 4)

    def com_mats(self):
        items = self.wbs_list.selection()
        
        if not items:
            return
        
        wbses = []
        for item in items:
            wbs = self.wbs_list.item(item, 'values')[1]
            wbses.append(wbs)
            
        his_display(self,'非标物料', wbses, 5)
        
    def do_popup(self, event):
        # display the popup menu
        try:
            self.popup.tk_popup(event.x_root, event.y_root, 0)
        finally:
            # make sure to release the grab (Tk 8.0a1 only)
            self.popup.grab_release() 
            
    def release_prj(self):
        if self.sel_index !=0:
            messagebox.showwarning('提示', '请在选项中选择已经Close但未release的项目')
            return
        
        sel_items = self.wbs_list.selection()
        error=[]
        for item in sel_items:
            i = project_cols.index('col2')
            wbs = self.wbs_list.item(item, 'values')[i]
            
            q = UnitInfo.update(status=6, modify_emp=login_info['uid'], modify_date=datetime.datetime.now()).where(UnitInfo.wbs_no==wbs)
            j= q.execute() 
            if j==0:
                error.append(wbs)
            else:
                self.wbs_dic.pop(wbs)
                self.wbs_list.delete(item)
                self.wbs_keys.remove(wbs)
        i_count=len(sel_items)
        i_fail =len(error)
        
        messagebox.showinfo('releas结果','release成功:'+str(i_count-i_fail)+';失败:'+str(i_fail))
        
    def select_func(self, event=None):
        self.__refresh_list()
    
    def refresh(self):
        self.pos=None
        i_index = self.func_list.index(self.project_func.get())
        self.refresh_prj(i_index) 
        self.update_data()
  
    def refresh_prj(self, i_sel):
        self.sel_index=i_sel
        for row in self.wbs_list.get_children():
            self.wbs_list.delete(row)
            
        if i_sel==0:
            reses = ProjectInfo.select(ProjectInfo, UnitInfo, ElevatorTypeDefine,LProcAct).join(UnitInfo, on=(ProjectInfo.project==UnitInfo.project)).switch(UnitInfo).\
                        join(ElevatorTypeDefine, on=(ElevatorTypeDefine.elevator_type_id==UnitInfo.elevator)).switch(UnitInfo).join(LProcAct, on=(UnitInfo.wbs_no==LProcAct.instance)).\
                        where((UnitInfo.status==5)&(LProcAct.action=='AT00000008')).order_by(UnitInfo.req_delivery_date.asc(), UnitInfo.wbs_no.asc()).naive()
        elif i_sel==1:
            reses = ProjectInfo.select(ProjectInfo, UnitInfo, ElevatorTypeDefine,LProcAct).join(UnitInfo, on=(ProjectInfo.project==UnitInfo.project)).switch(UnitInfo).\
                        join(ElevatorTypeDefine, on=(ElevatorTypeDefine.elevator_type_id==UnitInfo.elevator)).switch(UnitInfo).join(LProcAct, on=(UnitInfo.wbs_no==LProcAct.instance)).\
                        where((UnitInfo.status==6)&(LProcAct.action=='AT00000008')).order_by(UnitInfo.req_delivery_date.asc(), UnitInfo.wbs_no.asc()).naive() 
        elif i_sel==2:
            reses = ProjectInfo.select(ProjectInfo, UnitInfo, ElevatorTypeDefine).join(UnitInfo, on=(ProjectInfo.project==UnitInfo.project)).switch(UnitInfo).\
                        join(ElevatorTypeDefine, on=(ElevatorTypeDefine.elevator_type_id==UnitInfo.elevator)).\
                        where((UnitInfo.status==1)|(UnitInfo.status==3)).order_by(UnitInfo.req_delivery_date.asc(), UnitInfo.wbs_no.asc()).naive() 
        elif i_sel==3:
            reses = ProjectInfo.select(ProjectInfo, UnitInfo, ElevatorTypeDefine).join(UnitInfo, on=(ProjectInfo.project==UnitInfo.project)).switch(UnitInfo).\
                        join(ElevatorTypeDefine, on=(ElevatorTypeDefine.elevator_type_id==UnitInfo.elevator)).\
                        where(UnitInfo.status==3).order_by(UnitInfo.req_delivery_date.asc(), UnitInfo.wbs_no.asc()).naive() 
        elif i_sel==4:
            reses = ProjectInfo.select(ProjectInfo, UnitInfo, ElevatorTypeDefine).join(UnitInfo, on=(ProjectInfo.project==UnitInfo.project)).switch(UnitInfo).\
                        join(ElevatorTypeDefine, on=(ElevatorTypeDefine.elevator_type_id==UnitInfo.elevator)).\
                        where(UnitInfo.status==2).order_by(UnitInfo.req_delivery_date.asc(), UnitInfo.wbs_no.asc()).naive() 
        else:
            return False
        
        if not reses:
            return False
        
        self.wbs_dic={}
        self.wbs_keys=[]
        for r in reses:
            item=[]
            s_wbs = r.wbs_no
            
            self.wbs_keys.append(s_wbs)
            item.append(r.contract)
            item.append(r.wbs_no)
            item.append(r.project_name)
            item.append(r.lift_no)
            item.append(r.elevator)
            item.append(r.elevator_type)
            item.append(Status_Types[r.status])
            s_str = r.is_urgent and 'Y' or '' 
            item.append(s_str)
            item.append(Catalog_Types[r.project_catalog])
            s_str = none2str(date2str(r.req_delivery_date))
            item.append(s_str)
            s_str = none2str(date2str(r.req_configure_finish))
            item.append(s_str)
            if i_sel<2:
                item.append(r.finish_date)
            else:
                item.append('')

            item= item + self.get_nstd_info(s_wbs)
            item= item + self.get_mat_info(s_wbs)
            self.wbs_dic[s_wbs]=item
                                         
        return True
    
    def update_data(self): 
        i_status = project_cols.index('col7')
        i_urgent = project_cols.index('col8')
        i_catalog= project_cols.index('col9')
        for key in self.wbs_keys:
            s_status = self.wbs_dic[key][i_status]
            s_urgent = self.wbs_dic[key][i_urgent]
            s_catalog= self.wbs_dic[key][i_catalog]
            if s_status == Status_Types[4]:
                self.wbs_list.insert('', END, values=self.wbs_dic[key], tags=('stop',))
            elif s_urgent.upper()=='Y':
                self.wbs_list.insert('', END, values=self.wbs_dic[key], tags=('urgent',))
            elif s_catalog==Catalog_Types[5]:
                self.wbs_list.insert('', END, values=self.wbs_dic[key], tags=('pre-engineer',))
            elif s_catalog==Catalog_Types[6]:
                self.wbs_list.insert('', END, values=self.wbs_dic[key], tags=('lean',))
            else:
                self.wbs_list.insert('', END, values=self.wbs_dic[key])
        
        self.wbs_list.tag_configure('stop', background='red')
        self.wbs_list.tag_configure('urgent', background='green')
        self.wbs_list.tag_configure('pre-engineer', background='cyan')
        self.wbs_list.tag_configure('lean', background='yellow')
    
    def get_nstd_info(self, wbs):
        q_res = NonstdAppHeader.select(NonstdAppHeader.drawing_req_date, NonstdAppItemInstance.has_nonstd_draw, LProcAct.finish_date).join(NonstdAppItem, on=(NonstdAppHeader.nonstd==NonstdAppItem.nonstd)).switch(NonstdAppItem).join(NonstdAppItemInstance, on=(NonstdAppItem.index==NonstdAppItemInstance.index)).\
                    switch(NonstdAppItemInstance).join(LProcAct, on =(NonstdAppItemInstance.index_mat==LProcAct.instance)).where((NonstdAppItem.link_list.contains(wbs))&((LProcAct.action=='AT00000015')|(LProcAct.is_active==True))).order_by(LProcAct.finish_date.asc()).naive()
       
        if not q_res:
            return ['','','']
        
        item=[]
        b_nonstd=False
        dt_draw = str2datetime('1900-01-01 00:00:00')
        for r in q_res:
            b_temp = r.has_nonstd_draw
            dt_req = r.drawing_req_date
            
            if not b_temp:
                continue
            else:
                b_nonstd=b_temp
                dt_temp = r.finish_date
                    
            if b_nonstd and not dt_temp:
                dt_draw=dt_temp
                break
            
            if dt_draw < dt_temp:
                dt_draw=dt_temp
        
        s_str = b_nonstd and 'Y' or ''
        
        if dt_draw==str2datetime('1900-01-01 00:00:00'):
            dt_draw=None
        
        item.append(s_str)
        item.append(none2str(date2str(dt_req)))
        item.append(none2str(datetime2str(dt_draw)))
        
        return item             
                
    def get_mat_info(self, wbs):
        q_res = nstd_app_head.select(nstd_mat_fin.co_run_fin, nstd_mat_fin.co_run_fin_on, nstd_mat_table.mat_no).join(nstd_app_link).switch(nstd_app_head).join(nstd_mat_table).switch(nstd_mat_table).join(nstd_mat_fin).\
                        where((nstd_app_link.wbs_no==wbs)&(nstd_mat_fin.justify>=0)).naive()
        if not q_res:
            return ['','']
            
        item=[]
        dt_date = str2datetime('1900-01-01 00:00:00')
        for r in q_res:
            b_bom = r.co_run_fin
            
            if not b_bom:
                dt_date=None
                break
            else:
                dt_finish=r.co_run_fin_on
            
            if dt_date < dt_finish:
                dt_date = dt_finish
            
        s_str='Y'
        item.append(s_str)
        item.append(none2str(datetime2str(dt_date)))
        
        return item      
                               
    def __refresh_list(self):
        prj_thread= refresh_thread(self)
        prj_thread.setDaemon(True)
        prj_thread.start()   

    def __loop_refresh(self):
        self.__refresh_list()
        self.wbs_list.after(1800000, self.__loop_refresh)
        
class his_display(Toplevel):
    def __init__(self, parent, title, case_list, choice ):
        Toplevel.__init__(self, parent)

        self.withdraw()
        if parent.winfo_viewable():
            self.transient(parent)

        if title:
            self.title(title)

        self.parent = parent
        self.grab_set()
        self.geometry('800x600')
        
        self.create_widgets(choice)
        
        self.refresh_list(case_list, choice)
        
        self.protocol("WM_DELETE_WINDOW", self.close_wm)

        if self.parent is not None:
            center(self)

        self.deiconify() # become visible now

        # wait for window to appear on screen before calling grab_set
        self.wait_visibility()
        
        self.wait_window(self) 
        
    def create_widgets(self, choice):
        if choice==0:
            self.create_proc_tree()
        elif choice==1:
            self.create_nstd_app_tree()
        elif choice==2 or choice==3:
            self.create_nstd_design_tree()
        elif choice==4:
            self.create_nstd_content()
        elif choice==5:
            self.create_nstd_mat_content()
        ysb = ttk.Scrollbar(self, orient='vertical', command=self.proc_list.yview)
        xsb = ttk.Scrollbar(self, orient='horizontal', command=self.proc_list.xview)  
        self.proc_list.grid(row=0, column=0, rowspan=6, columnspan=4, sticky='nsew') 
        xsb.grid(row=6, column=0, columnspan=4, sticky='ew')
        ysb.grid(row=0, column=4, rowspan=6, sticky='ns')
        
        self.close_button= Button(self, text='退出')
        self.close_button.grid(row=7, column=2, columnspan=1, sticky='nsew')
        self.close_button['command']=self.close_wm
        self.export_button = Button(self, text='导出')
        self.export_button.grid(row=7, column=1, sticky='nsew')
        self.export_button['command']=self.export_excel
        if choice!=5:
            self.export_button.grid_forget()
        self.columnconfigure(3, weight=1)
        self.rowconfigure(5, weight=1)
        
        self.bind('<Escape>', self.close_wm)
        
    def export_excel(self):
        items = self.proc_list.get_children('')
        if not items:
            return

        file_str=filedialog.asksaveasfilename(title="导出文件", filetypes=[('excel file','.xlsx')])
        if not file_str:
            return

        if not file_str.endswith(".xlsx"):
            file_str+=".xlsx"

        wb=Workbook()
        ws=wb.worksheets[0]
        ws.title='物料清单'
        col_size = len(self.cols_his)
        for i in range(col_size):
            ws.cell(row=1,column=i+1).value=self.tree_head_his[i]  
        n=0 
        
        dic_str = {}
        for item in items:
            if self.proc_list.parent(item) !='':
                continue
            
            for i in range(col_size):
                ws.cell(row=n+2, column=i+1).value = self.proc_list.item(item, 'values')[i]
                
            n+=1

        if excel_xlsx.save_workbook(workbook=wb, filename=file_str):
            messagebox.showinfo("输出","成功输出!")         
        
    def close_wm(self, event=None):
        self.withdraw()
        self.update_idletasks() 
        Toplevel.destroy(self)

    def create_proc_tree(self):
        self.proc_list = ttk.Treeview(self, show='headings', columns=['col1','col2','col3','col4','col5','col6','col7','col8','col9','col10'])
        self.proc_list.heading('col1', text='Unit WBS NO.')
        self.proc_list.heading('col2', text='项目信息')
        self.proc_list.heading('col3', text='梯号')
        self.proc_list.heading('col4', text='流程')
        self.proc_list.heading('col5', text='步骤')
        self.proc_list.heading('col6', text='负责人')
        self.proc_list.heading('col7', text='收到日期')
        self.proc_list.heading('col8', text='完成日期')
        self.proc_list.heading('col9', text='配置完成日期')
        self.proc_list.heading('col10', text='发货期')        
        self.proc_list.column('col1', width=100, anchor='w')
        self.proc_list.column('col2', width=100, anchor='w')
        self.proc_list.column('col3', width=50, anchor='w')
        self.proc_list.column('col4', width=100, anchor='w')
        self.proc_list.column('col5', width=100, anchor='w')
        self.proc_list.column('col6', width=100, anchor='w')
        self.proc_list.column('col7', width=100, anchor='w')
        self.proc_list.column('col8', width=100, anchor='w')
        self.proc_list.column('col9', width=100, anchor='w')
        self.proc_list.column('col10', width=100, anchor='w')

    def create_nstd_app_tree(self):
        self.proc_list = ttk.Treeview(self, show='headings', columns=['col1','col2','col3','col4','col5','col6','col7','col8','col9'])
        self.proc_list.heading('col1', text='非标编号')
        self.proc_list.heading('col2', text='项目信息')
        self.proc_list.heading('col3', text='流程')
        self.proc_list.heading('col4', text='步骤')
        self.proc_list.heading('col5', text='负责人')
        self.proc_list.heading('col6', text='收到日期')
        self.proc_list.heading('col7', text='完成日期')
        self.proc_list.heading('col8', text='配置完成日期')
        self.proc_list.heading('col9', text='发货期')        
        self.proc_list.column('col1', width=100, anchor='w')
        self.proc_list.column('col2', width=100, anchor='w')
        self.proc_list.column('col3', width=100, anchor='w')
        self.proc_list.column('col4', width=100, anchor='w')
        self.proc_list.column('col5', width=100, anchor='w')
        self.proc_list.column('col6', width=100, anchor='w')
        self.proc_list.column('col7', width=100, anchor='w')
        self.proc_list.column('col8', width=100, anchor='w')
        self.proc_list.column('col9', width=100, anchor='w')
    
    def create_nstd_design_tree(self):
        self.proc_list = ttk.Treeview(self, show='headings', columns=['col1','col2','col3','col4','col5','col6','col7','col8','col9','col10'])
        self.proc_list.heading('col1', text='非标编号')
        self.proc_list.heading('col2', text='非标物料申请号')
        self.proc_list.heading('col3', text='项目信息')
        self.proc_list.heading('col4', text='流程')
        self.proc_list.heading('col5', text='步骤')
        self.proc_list.heading('col6', text='负责人')
        self.proc_list.heading('col7', text='收到日期')
        self.proc_list.heading('col8', text='完成日期')
        self.proc_list.heading('col9', text='配置完成日期')
        self.proc_list.heading('col10', text='发货期')        
        self.proc_list.column('col1', width=100, anchor='w')
        self.proc_list.column('col2', width=100, anchor='w')
        self.proc_list.column('col3', width=100, anchor='w')
        self.proc_list.column('col4', width=100, anchor='w')
        self.proc_list.column('col5', width=100, anchor='w')
        self.proc_list.column('col6', width=100, anchor='w')
        self.proc_list.column('col7', width=100, anchor='w')
        self.proc_list.column('col8', width=100, anchor='w')
        self.proc_list.column('col9', width=100, anchor='w')
        self.proc_list.column('col10', width=100, anchor='w') 
        
    def create_nstd_content(self):
        self.proc_list = ttk.Treeview(self, show='headings', columns=['col1','col2','col3','col4','col5','col6','col7','col8','col9','col10','col11','col12','col13'])
        self.proc_list.heading('col1', text='非标任务编号')
        self.proc_list.heading('col2', text='非标物料申请号')
        self.proc_list.heading('col3', text='项目信息')
        self.proc_list.heading('col4', text='物料需求日')
        self.proc_list.heading('col5', text='图纸需求日')
        self.proc_list.heading('col6', text='非标类别')
        self.proc_list.heading('col7', text='非标原因')
        self.proc_list.heading('col8', text='分值分布')
        self.proc_list.heading('col9', text='非标负责人')
        self.proc_list.heading('col10', text='分任务描述') 
        self.proc_list.heading('col11', text='分任务非标工程师')
        self.proc_list.heading('col12', text='分任务状态')
        self.proc_list.heading('col13', text='关联梯')
        self.proc_list.column('col1', width=80, anchor='w')
        self.proc_list.column('col2', width=80, anchor='w')
        self.proc_list.column('col3', width=100, anchor='w')
        self.proc_list.column('col4', width=100, anchor='w')
        self.proc_list.column('col5', width=100, anchor='w')
        self.proc_list.column('col6', width=100, anchor='w')
        self.proc_list.column('col7', width=100, anchor='w')
        self.proc_list.column('col8', width=100, anchor='w')
        self.proc_list.column('col9', width=100, anchor='w')
        self.proc_list.column('col10', width=100, anchor='w')  
        self.proc_list.column('col11', width=100, anchor='w') 
        self.proc_list.column('col12', width=100, anchor='w') 
        self.proc_list.column('col13', width=100, anchor='w') 
        
    def create_nstd_mat_content(self):
        self.cols_his = ['col0','col1','col2','col3','col4','col5','col6','col7','col8','col9','col10','col11','col12','col13','col14','col15','col16','col17','col18','col19','col20']
        self.tree_head_his=['导入日期','非标编号','判断','物料号','物料名称(中)','物料名称(英)','图号','单位','备注','RP','BoxId','申请人','自制完成','负责人','完成日期','价格维护','负责人','完成日期','CO-run','负责人','完成日期']                          
        self.proc_list = ttk.Treeview(self, show='headings',columns=self.cols_his)
        for col in self.cols_his:
            i = self.cols_his.index(col)
            self.proc_list.heading(col, text=self.tree_head_his[i])  
        
        self.proc_list.column('col0', width=80, anchor='w')
        self.proc_list.column('col1', width=100, anchor='w')
        self.proc_list.column('col2', width=80, anchor='w')
        self.proc_list.column('col3', width=80, anchor='w')
        self.proc_list.column('col4', width=100, anchor='w')
        self.proc_list.column('col5', width=100, anchor='w')
        self.proc_list.column('col6', width=100, anchor='w')
        self.proc_list.column('col7', width=40, anchor='w')
        self.proc_list.column('col8', width=200, anchor='w')
        self.proc_list.column('col9', width=50, anchor='w')
        self.proc_list.column('col10', width=50, anchor='w')
        self.proc_list.column('col11', width=60, anchor='w')
        self.proc_list.column('col12', width=60, anchor='w')
        self.proc_list.column('col13', width=50, anchor='w')
        self.proc_list.column('col14', width=100, anchor='w')
        self.proc_list.column('col15', width=60, anchor='w')
        self.proc_list.column('col16', width=50, anchor='w')
        self.proc_list.column('col17', width=100, anchor='w')
        self.proc_list.column('col18', width=60, anchor='w')
        self.proc_list.column('col19', width=50, anchor='w')
        self.proc_list.column('col20', width=100, anchor='w')

    def refresh_list(self, cases, choice):
        for item in self.proc_list.get_children():
            self.proc_list.delete(item)
            
        if choice==0:
            q_res = LProcAct.select(LProcAct, UnitInfo,ProjectInfo, SEmployee.name,WorkflowInfo.workflow_name, ActionInfo.action_name).join(UnitInfo, on=(LProcAct.instance==UnitInfo.wbs_no)).\
                            switch(UnitInfo).join(ProjectInfo, on=(ProjectInfo.project==UnitInfo.project)).switch(LProcAct).join(WorkflowInfo, on=(WorkflowInfo.workflow==LProcAct.workflow)).join(ActionInfo, on=(ActionInfo.action==LProcAct.action)).join(SEmployee, on=(LProcAct.operator==SEmployee.employee)).\
                            where(((LProcAct.workflow=='WF0002')|(LProcAct.workflow=='WF0006'))& (LProcAct.instance.in_(cases))).order_by(UnitInfo.wbs_no.asc(),LProcAct.flow_ser.asc()).naive()
        elif choice==1:
            indexes=[]
            for r in cases:
                index_res = NonstdAppItem.select(NonstdAppItem.index).where(NonstdAppItem.link_list.contains(r))
                if not index_res:
                    continue
                for i in index_res:
                    if i.index not in indexes:
                        indexes.append(i.index)
                    else:
                        continue
                        
            if len(indexes)==0:
                return
            
            q_res = LProcAct.select(LProcAct, ProjectInfo,WorkflowInfo, ActionInfo, SEmployee.name,NonstdAppHeader.mat_req_date, NonstdAppHeader.drawing_req_date).join(NonstdAppItem, on=(NonstdAppItem.index==LProcAct.instance)).switch(NonstdAppItem).join(NonstdAppHeader, on=(NonstdAppItem.nonstd==NonstdAppHeader.nonstd)).\
                        switch(NonstdAppHeader).join(ProjectInfo, on=(NonstdAppHeader.project==ProjectInfo.project)).switch(LProcAct).join(WorkflowInfo, on=(WorkflowInfo.workflow==LProcAct.workflow)).join(ActionInfo, on=(ActionInfo.action==LProcAct.action)).join(SEmployee, on=(LProcAct.operator==SEmployee.employee)).\
                        where((LProcAct.instance.in_(indexes))&(NonstdAppHeader.status>=0)).order_by(LProcAct.instance.asc(), LProcAct.flow_ser.asc()).naive()
        elif choice==2:
            indexes=[]
            for r in cases:
                index_res = NonstdAppItemInstance.select(NonstdAppItemInstance.index_mat).join(NonstdAppItem, on = (NonstdAppItemInstance.index==NonstdAppItem.index)).where(NonstdAppItem.link_list.contains(r))
                if not index_res:
                    continue
                for i in index_res:
                    if i.index_mat not in indexes:
                        indexes.append(i.index_mat)
                    else:
                        continue
            if len(indexes)==0:
                return
            
            q_res = LProcAct.select(LProcAct, ProjectInfo,WorkflowInfo, ActionInfo, SEmployee.name, NonstdAppItemInstance.nstd_mat_app,NonstdAppHeader.mat_req_date, NonstdAppHeader.drawing_req_date).\
                        join(NonstdAppItemInstance, on=(NonstdAppItemInstance.index_mat==LProcAct.instance)).switch(NonstdAppItemInstance).join(NonstdAppItem, on=(NonstdAppItemInstance.index==NonstdAppItem.index)).switch(NonstdAppItem).join(NonstdAppHeader, on=(NonstdAppHeader.nonstd==NonstdAppItem.nonstd)).switch(NonstdAppHeader).\
                        join(ProjectInfo, on=(NonstdAppHeader.project==ProjectInfo.project)).switch(LProcAct).join(WorkflowInfo, on=(WorkflowInfo.workflow==LProcAct.workflow)).join(ActionInfo, on=(ActionInfo.action==LProcAct.action)).join(SEmployee, on=(LProcAct.operator==SEmployee.employee)).\
                        where((LProcAct.instance.in_(indexes))&(NonstdAppItemInstance.status>=0)&(LProcAct.workflow=='WF0005')).order_by(LProcAct.instance.asc(), LProcAct.flow_ser.asc()).naive()
        elif choice==3:
            indexes=[]
            for r in cases:
                index_res = NonstdAppItemInstance.select(NonstdAppItemInstance.index_mat).join(NonstdAppItem, on = (NonstdAppItemInstance.index==NonstdAppItem.index)).where(NonstdAppItem.link_list.contains(r))
                if not index_res:
                    continue
                for i in index_res:
                    if i.index_mat not in indexes:
                        indexes.append(i.index_mat)
                    else:
                        continue
            if len(indexes)==0:
                return
            
            q_res = LProcAct.select(LProcAct,ProjectInfo,WorkflowInfo,ActionInfo,NonstdAppItemInstance.nstd_mat_app,SEmployee.name,NonstdAppHeader.mat_req_date, NonstdAppHeader.drawing_req_date).\
                        join(NonstdAppItemInstance, on=(NonstdAppItemInstance.index_mat==LProcAct.instance)).switch(NonstdAppItemInstance).join(NonstdAppItem, on=(NonstdAppItemInstance.index==NonstdAppItem.index)).switch(NonstdAppItem).join(NonstdAppHeader, on=(NonstdAppHeader.nonstd==NonstdAppItem.nonstd)).switch(NonstdAppHeader).\
                        join(ProjectInfo, on=(NonstdAppHeader.project==ProjectInfo.project)).switch(LProcAct).join(WorkflowInfo, on=(WorkflowInfo.workflow==LProcAct.workflow)).join(ActionInfo, on=(ActionInfo.action==LProcAct.action)).join(SEmployee, on=(LProcAct.operator==SEmployee.employee)).\
                        where((LProcAct.instance.in_(indexes))&(NonstdAppItemInstance.status>=0)&(LProcAct.workflow=='WF0004')).order_by(LProcAct.instance.asc(), LProcAct.flow_ser.asc()).naive()
        elif choice==4:
            indexes=[]
            for r in cases:
                index_res = NonstdAppItemInstance.select(NonstdAppItemInstance.index_mat).join(NonstdAppItem, on = (NonstdAppItemInstance.index==NonstdAppItem.index)).where(NonstdAppItem.link_list.contains(r))
                if not index_res:
                    continue
                for i in index_res:
                    if i.index_mat not in indexes:
                        indexes.append(i.index_mat)
                    else:
                        continue 
                        
            if len(indexes)==0:
                return
            
            q_res = NonstdAppItemInstance.select(NonstdAppItemInstance, NonstdAppItem, NonstdAppHeader.mat_req_date, NonstdAppHeader.drawing_req_date, ProjectInfo, SEmployee.name).join(NonstdAppItem, on=(NonstdAppItemInstance.index==NonstdAppItem.index)).switch(NonstdAppItem).join(NonstdAppHeader, on=(NonstdAppHeader.nonstd==NonstdAppItem.nonstd)).switch(NonstdAppHeader).\
                       join(ProjectInfo, on=(ProjectInfo.project==NonstdAppHeader.project)).switch(NonstdAppItemInstance).join(SEmployee, on=(NonstdAppItemInstance.res_engineer==SEmployee.employee)).\
                       where((NonstdAppItemInstance.index_mat.in_(indexes))&(NonstdAppItemInstance.status>=0)).order_by(NonstdAppItemInstance.index_mat.asc()).naive()
        elif choice==5:
            q_res = nstd_app_head.select(nstd_app_head, nstd_mat_table, nstd_mat_fin).join(nstd_app_link).switch(nstd_app_head).join(nstd_mat_table).switch(nstd_mat_table).join(nstd_mat_fin).\
                        where((nstd_app_link.wbs_no.in_(cases))&(nstd_mat_fin.justify>=0)).naive()
        self.insert_data(q_res, choice)
        
    def insert_data(self, q_res, choice):
        if not q_res:
            return
        init_str=''
        b_switch=False
        if choice==0: 
            for r in q_res:
                item=[]
                if len(init_str)!=0 and init_str!=r.instance:
                    b_switch = not b_switch
                init_str=r.instance
                item.append(r.instance)
                item.append(r.contract+' '+r.project_name)
                item.append(r.lift_no)
                item.append(r.workflow_name)
                item.append(r.action_name)
                item.append(r.name)
                item.append(none2str(datetime2str(r.start_date)))
                b_active = r.is_active
                if not b_active:
                    item.append(none2str(datetime2str(r.finish_date)))
                else:
                    item.append('')
                item.append(none2str(date2str(r.req_configure_finish)))
                item.append(none2str(date2str(r.req_delivery_date)))
                if b_switch:
                    self.proc_list.insert('', END, values=item, tags=('switch',))
                elif not b_active:
                    self.proc_list.insert('', END, values=item, tags=('unactive',))
                else:
                    self.proc_list.insert('', END, values=item, tags=('active',))
        elif choice==1:
            for r in q_res:
                item=[]
                if len(init_str)!=0 and init_str!=r.instance:
                    b_switch = not b_switch
                init_str=r.instance
                item.append(r.instance)
                item.append(r.contract+' '+r.project_name+' '+r.project)
                item.append(r.workflow_name)
                item.append(r.action_name)
                item.append(r.name)
                item.append(none2str(datetime2str(r.start_date)))
                b_active=r.is_active
                if not b_active:
                    item.append(none2str(datetime2str(r.finish_date)))
                else:
                    item.append('')
                item.append(none2str(date2str(r.mat_req_date)))
                item.append(none2str(date2str(r.drawing_req_date)))
                if b_switch:
                    self.proc_list.insert('', END, values=item, tags=('switch',))
                elif not b_active:
                    self.proc_list.insert('', END, values=item, tags=('unactive',))
                else:
                    self.proc_list.insert('', END, values=item, tags=('active',))
        elif choice==2 or choice==3:
            for r in q_res:
                item=[]
                if len(init_str)!=0 and init_str!=r.instance:
                    b_switch = not b_switch
                init_str=r.instance
                item.append(r.instance)
                item.append(r.nstd_mat_app)
                item.append(r.contract+' '+r.project_name+' '+r.project)
                item.append(r.workflow_name)
                item.append(r.action_name)
                item.append(r.name)
                item.append(none2str(datetime2str(r.start_date)))
                b_active=r.is_active
                if not b_active:
                    item.append(none2str(datetime2str(r.finish_date)))
                else:
                    item.append('')
                item.append(none2str(date2str(r.mat_req_date)))
                item.append(none2str(date2str(r.drawing_req_date)))
                if b_switch:
                    self.proc_list.insert('', END, values=item, tags=('switch',))
                elif not b_active:
                    self.proc_list.insert('', END, values=item, tags=('unactive',))
                else:
                    self.proc_list.insert('', END, values=item, tags=('active',))
        elif choice==4:
            for r in q_res:
                item=[]
                item.append(r.index_mat)
                item.append(r.nstd_mat_app)
                item.append(r.contract+' '+r.project_name+' '+r.project)
                item.append(none2str(date2str(r.mat_req_date)))
                item.append(none2str(date2str(r.drawing_req_date)))
                item.append(none2str(r.nonstd_catalog))
                item.append(none2str(r.nonstd_desc))
                item.append(none2str(r.nonstd_value))
                s_person=r.res_person
                try:
                    s_res_person= SEmployee.get(SEmployee.employee==s_person).name
                except SEmployee.DoesNotExist:
                    s_res_person=s_person
                item.append(s_res_person)
                item.append(none2str(r.instance_nstd_desc))
                item.append(none2str(r.name))
                item.append(Status_Types[r.status])
                item.append(none2str(r.link_list))
                self.proc_list.insert('', END, values=item)
        elif choice==5:
            for r in q_res:
                item=[]
                nstd_app_id = r.nstd_app
                index_mat_id = r.index_mat
                item.append(none2str(date2str(r.modify_on)))
                item.append(nstd_app_id)
                item.append(Justify_Types[r.justify])
                item.append(r.mat_no)
                item.append(r.mat_name_cn)
                item.append(r.mat_name_en)
                item.append(none2str(r.drawing_no))
                item.append(r.mat_unit)
                item.append(none2str(r.comments))
                item.append(none2str(r.rp))
                item.append(none2str(r.box_code_sj))
                app_per = r.app_person
                if app_per is None:
                    app_per = r.mat_app_person
                elif len(app_per)==0:
                    app_per = r.mat_app_person
            
                item.append(app_per)
                temp = (r.mbom_fin and 'Y' or '')
                item.append(temp)
                if temp.upper()=='Y':
                    item.append(get_name(r.mbom_fin_by))
                    item.append(none2str(r.mbom_fin_on))
                else:
                    item.append('')
                    item.append('')
                    temp=(r.pu_price_fin and 'Y' or '')
                    item.append(temp)
                if temp.upper()=='Y':
                    item.append(get_name(r.pu_price_fin_by))
                    item.append(none2str(r.pu_price_fin_on))
                else:
                    item.append('')
                    item.append('')                
                temp = (r.co_run_fin and 'Y' or '')
                item.append(temp)
                if temp.upper()=='Y':
                    item.append(get_name(r.co_run_fin_by))
                    item.append(none2str(r.co_run_fin_on))
                else:
                    item.append('')
                    item.append('')
                    
                if temp.upper()=='Y':
                    self.proc_list.insert('', END, values=item, tags=('unactive',))
                else:
                    self.proc_list.insert('', END, values=item, tags=('active',))                
                                   
        self.proc_list.tag_configure('unactive', background='lightgrey')
        self.proc_list.tag_configure('active', background='lightpink')
        self.proc_list.tag_configure('switch', background='white')
                