#coding=utf-8
'''
Created on 2017年1月24日

@author: 10256603
'''
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


def dic_to_list(m_dict):
    m_list = []
      
    for c in col_header:
        if m_dict[c] is None:
            m_list.append('')
        else:
            m_list.append(m_dict[c])
        
    return m_list
          
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
        self.widgets_group1.grid(row=0, column=0, columnspan=6, sticky=NSEW)
        self.im_button = Button(self.widgets_group1, text='导入数据')
        self.im_button.pack(side='left')
        self.im_button['command']=self.import_data
        self.im_cs_button = Button(self.widgets_group1, text='导入CS数据')
        self.im_cs_button.pack(side='left')
        self.im_cs_button['command']=self.import_data
        self.eds_active_button = Button(self.widgets_group1, text='激活EDS非标申请')
        self.eds_active_button.pack(side='left')
        self.eds_active_button['command']= self.eds_active
        self.out_button = Button(self.widgets_group1, text='导出PDM导入表')
        self.out_button.pack(side='left')
        self.out_button['command']=self.out_pdm_excel
        self.out_active_button = Button(self.widgets_group1, text='导出物料申请表')
        self.out_active_button.pack(side='left')
        self.out_active_button['command'] = self.out_active_excel
        self.del_button = Button(self.widgets_group1, text='删除物料(by申请号)')
        self.del_button.pack(side='left')
        self.del_button['command']=self.del_mats
        if login_info['perm'][0]!='2' and login_info['perm'][0]!='9':
            self.widgets_group1.grid_forget()  

        self.widgets_group2=Frame(self)
        self.widgets_group2.grid(row=1,column=0, columnspan=7, sticky=NSEW) 
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
        self.fin_button = Button(self.widgets_group2, text='直接关闭物料')
        self.fin_button.pack(side='left')
        self.fin_button['command']=self.mat_close
        if login_info['perm'][0]!='3' and login_info['perm'][0]!='9':
            self.widgets_group2.grid_forget()

        self.wl_str = StringVar()
        self.wl_entry = Entry(self, textvariable=self.wl_str)
        self.wl_entry.grid(row=0, column=6, sticky = EW)
        self.wl_entry.bind("<Return>", self.wl_result)

        table_panel = Frame(self)
        cols = ['col0','col1','col2','col3','col4','col5','col6','col7','col8','col9','col10','col11']
        self.mat_table = ttk.Treeview(table_panel, show='headings',columns=cols, selectmode='extended')
        style = ttk.Style()
        style.configure("Treeview", font=('TkDefaultFont', 10))
        style.configure("Treeview.Heading", font=('TkDefaultFont', 9)) 
        self.mat_table.heading("#0",text='')       
        for col in cols:
            i = cols.index(col)
            #self.mat_table.heading(col, text=tree_head[i])
            self.mat_table.heading(col, text=col_header[i] , command=lambda _col=col: treeview_sort_column(self.mat_table, _col, False))
            
        #self.mat_table.column('#0', width=20)
        self.mat_table.column('col0', width=100, anchor='w')
        self.mat_table.column('col1', width=80, anchor='w')
        self.mat_table.column('col2', width=100, anchor='w')
        self.mat_table.column('col3', width=150, anchor='w')
        self.mat_table.column('col4', width=150, anchor='w')
        self.mat_table.column('col5', width=100, anchor='w')
        self.mat_table.column('col6', width=80, anchor='w')
        self.mat_table.column('col7', width=200, anchor='w')
        self.mat_table.column('col8', width=100, anchor='w')
        self.mat_table.column('col9', width=100, anchor='w')
        self.mat_table.column('col10', width=100, anchor='w')
        self.mat_table.column('col11', width=100, anchor='w')
        
        ysb = ttk.Scrollbar(table_panel, orient='vertical', command=self.mat_table.yview)
        xsb = ttk.Scrollbar(table_panel, orient='horizontal', command=self.mat_table.xview)  
        self.mat_table.grid(row=0, column=0, rowspan=3, columnspan=3, sticky='nsew')
        
        ysb.grid(row=0, column=3,rowspan=3, sticky='ns')
        xsb.grid(row=3, column=0, columnspan=3, sticky='ew')    

        table_panel.grid(row=2, column=0, rowspan=6, columnspan=10, sticky=NSEW )
        table_panel.columnconfigure(1, weight=1)
        table_panel.rowconfigure(1, weight=1)
        
        self.grid(row=0, column=0, sticky=NSEW)
        self.columnconfigure(9, weight=1)
        self.rowconfigure(5, weight=1)  
    
    def eds_active(self):
        nstd_id = simpledialog.askstring('非标申请编号','请输入完整非标申请编号(不区分大小写):')
        if not nstd_id:
            return
        
        nstd_id = nstd_id.upper().strip()
        
        if len(nstd_id)==0:
            return
        
        res = nstd_app_head.select(nstd_mat_table.mat_no).join(nstd_mat_table).where(nstd_app_head.nstd_app==nstd_id).naive()
        
        if not res:
            messagebox.showwarning('警告','数据库中无与非标申请号相关的物料。')
            return
        
        mats=[]
        for r in res:
            mat = none2str(r.mat_no)
            mats.append(mat) 
            
            try:
                rs=nstd_mat_fin.get(nstd_mat_fin.mat_no==mat)
                if rs.justify==-1:
                    change_log('nstd_mat_fin', 'justify', mat, '-1', '0')                          
            except nstd_mat_fin.DoesNotExist:
                continue
            
        q = nstd_mat_fin.update(justify=0).where(nstd_mat_fin.mat_no << mats and nstd_mat_fin.justify==-1)
        
        i = q.execute()
        
        messagebox.showinfo('更新完成', '共计'+str(i)+'个物料激活!')
    
    def out_active_excel(self):  
        if self.mr_rows ==0:
            messagebox.showwarning('提示', '请先输入非标申请号，得到物料清单号再点击输出.')
            return
        
        file_str=filedialog.asksaveasfilename(title="导出文件", initialfile='temp',filetypes=[('excel file','.xlsx')])
        if not file_str:
            return

        if not file_str.endswith(".xlsx"):
            file_str+=".xlsx"

        temp_file = os.path.join(cur_dir(),'template.xlsx')
        wb = load_workbook(temp_file)
        ws = wb.active
        nstd =''

        for i in range(1, self.mr_rows+1):
            ws.cell(row=i+3, column=1).value = date2str(datetime.datetime.today())
            ws.cell(row=i+3, column=7).value = self.mat_dic[i][col_header[2]]
            ws.cell(row=i+3, column=9).value = self.mat_dic[i][col_header[10]]
            ws.cell(row=i+3, column=10).value = self.mat_dic[i][col_header[3]]
            ws.cell(row=i+3, column=14).value = self.mat_dic[i][col_header[5]]
            if len(self.mat_dic[i][col_header[5]])!=0 and self.mat_dic[i][col_header[5]].lower()!='no':
                ws.cell(row=i+3, column=13).value='是'
            ws.cell(row=i+3, column=15).value = self.mat_dic[i][col_header[9]]
            ws.cell(row=i+3, column=17).value = self.mat_dic[i][col_header[8]]
            ws.cell(row=i+3, column=18).value = self.mat_dic[i][col_header[4]]
            ws.cell(row=i+3, column=19).value = self.mat_dic[i][col_header[6]]
            ws.cell(row=i+3, column=12).value = self.mat_dic[i][col_header[7]]
            
            if nstd != self.mat_dic[i][col_header[0]] or  len(nstd)==0:             
                nstd=self.mat_dic[i][col_header[0]]              
                
            self.fill_others(ws,i+3, nstd)   
                
            ws.cell(row=i+3, column=23).value = nstd
            
        if excel_xlsx.save_workbook(workbook=wb, filename=file_str):
            messagebox.showinfo("输出","成功输出!")
            
    def fill_others(self,ws, row, nstd): 
        try:
            res= nstd_app_head.select(nstd_app_link.wbs_no).join(nstd_app_link).where(nstd_app_head.nstd_app==nstd).naive()
        except nstd_app_head.DoesNotExist:
            return
        
        wbses =[]
        for r in res:
            wbs = r.wbs_no
            wbses.append(wbs)
            
        if len(wbses)==0:
            return
        
        try:   
            r =  ProjectInfo.select(ProjectInfo.project_name,UnitInfo.wbs_no, ElevatorTypeDefine.elevator_type).join(UnitInfo,on=(ProjectInfo.project==UnitInfo.project)).switch(UnitInfo)\
                .join(ElevatorTypeDefine).where(UnitInfo.wbs_no==wbses[0]).naive().get()
        except ProjectInfo.DoesNotExist:
            return
        
        ws.cell(row=row,column=5).value = r.project_name
        ws.cell(row=row, column=4).value = self.combine_wbs(wbses)
        ws.cell(row=row, column=6).value = r.elevator_type
        
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

    def out_pdm_excel(self):
        if self.mr_rows ==0:
            messagebox.showwarning('提示', '请先输入非标申请号，得到物料清单号再点击输出.')
            return
        
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

    def import_data(self):
        file_list = filedialog.askopenfilenames(title="导入文件", filetypes=[('excel file','.xlsx'),('excel file','.xlsm')])
        if not file_list:
            return

        self.mr_rows = 0
        self.mat_dic={}
        self.nstd_app_id = []
        for file in file_list:
            self.read_excel_files(file)
            
        self.refresh_table()
        
    def refresh_table(self):
        for row in self.mat_table.get_children():
            self.mat_table.delete(row)
        
        mats_length = len(self.mat_dic)   
        for i in range(1, mats_length+1):
            mat_value = dic_to_list(self.mat_dic[i])
            
            self.mat_table.insert('', END, values=mat_value)
                    
    def read_excel_files(self, file):
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
            
            try:
                r_str = ws1.cell(row=rx+4, column=7).value
            except:
                break
            
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
            
            if len(s_str)==0:
                pass
            elif not s_str.startswith('E/'):
                s_str=s_str.replace('E','E/')
            
                
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
                
                if len(item['wbs'])!=0 and item['wbs'].startswith('E/'):
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
               
    def  read_workbook(self, wb, sheetname):
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
            nstd_id=nstd_id[:i_pos+2]
        
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
            
            if not wbs.startswith('E/'):
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

        if r>0:
            messagebox.showinfo('提示','删除成功')

    def wl_result(self, event):
        self.mr_rows = 0
        self.mat_dic = {}
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
    
        self.refresh_table()

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
        
    def mat_close(self):
        d = ask_list('物料拷贝器')
        if not d:
            return
                
        comment = simpledialog.askstring('原因','此操作不可逆，请谨慎！\n输入关闭原因:')
        if not comment:
            return
        
        i_error=0
        i_succ=0
        error_list=''
        for mat in d:
            try:
                r=nstd_mat_fin.get(nstd_mat_fin.mat_no==mat)
            except nstd_mat_fin.DoesNotExist:
                i_error+=1
                error_list=error_list+mat+';'
                continue
            
            if r.justify>0:    
                s = nstd_mat_fin.update(co_run_fin=True,co_run_fin_on=datetime.datetime.now(), co_run_fin_by=login_info['uid'], co_run_fin_remark=none2str(comment)).where(nstd_mat_fin.mat_no==mat)
            else:
                s = nstd_mat_fin.update(justify=13, co_run_fin=True,co_run_fin_on=datetime.datetime.now(), co_run_fin_by=login_info['uid'], co_run_fin_remark=none2str(comment)).where(nstd_mat_fin.mat_no==mat)
            if s.execute()==0:
                i_error+=1
                error_list=error_list+mat+';'
            else:
                i_succ+=1
        
        messagebox.showinfo('结果', '处理完成('+str(i_error+i_succ)+'):成功('+str(i_succ)+')!失败('+str(i_error)+'):'+error_list+'!')
            
    def set_justify(self, choice, mats, comment=None):        
        self.mr_rows=0
        self.mat_dic={}

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
                if r.app_person is None or len(r.app_person)==0 or r.app_person=='\'\'':
                    mr_line[col_header[10]]=r.mat_app_person
                else:
                    mr_line[col_header[10]]=r.app_person
                self.mat_dic[self.mr_rows+1]=mr_line
            self.mr_rows+=1 
            
        self.refresh_table()                   

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