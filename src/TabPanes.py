#!/usr/bin/env python3
#coding:utf-8
"""
  Author:   10256603<mikewolf.li@tkeap.com>
  Purpose: 
  Created: 2016/4/7
"""
from global_list import *
from mbom_dataset import UnitInfo, ProjectInfo
global login_info

header_line1 = ['TDM_ID','CN_DRAWING_NUMBER','CN_OLD_DRAWING_NUMBER','CN_PART_NAME_CHINESE','TDM_DESCRIPTION',
                'CN_OUTLINE_SIZE','CN_PART_TYPE|TN_PART_TYPE|TDM_NAME','CN_MATERIAL_LOC','CN_PART_WEIGHT','TDM_UOM|TDM_UNIT_OF_MEASURE|TDM_NAME',
                'CN_COMMENT','CN_COMPONENT_GROUP_CODE_SJ|TN_SYSTEMS|Cn_code_value','CN_COMPONENT_GROUP_CODE_ZS|TN_SYSTEMS|Cn_code_value',
                'CN_ACTUAL_BOX_SJ|TN_SYSTEMS|Cn_code_value','CN_ACTUAL_BOX_ZS|TN_SYSTEMS|Cn_code_value','CN_IN_EO_NO']
header_line2 = ['ID--Part','Drawing Number--Part','Old Drawing Number--Part','Part Name Chinese--Part','Description--Part','Outline Size--Part',
                'Item Type--Part','Material Loc--Part','Item Weight--Part','Units of Measure--Part','Comment--Part']

col_header = ['非标物料申请编号','判断','物料号','物料名称(中)','物料名称(英)','图号','单位','备注','RP','BoxId','申请人','旧物料号']
mat_db_header =['nstd_app','mat_no','mat_name_cn','mat_name_en','drawing_no','mat_unit','comments','rp','box_code_sj','justify','app_person','old_mat_no']

def change_log(table,section,key, old,new):
    q = s_change_log.insert(table_name=table,change_section=section,key_word=str(key),old_value=str(old),new_value=str(new),log_on=datetime.datetime.now(), log_by=login_info['uid'] )
    q.execute()
        
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
        self.fin_button = Button(self.widgets_group2, text='直接关闭物料')
        self.fin_button.pack(side='left')
        self.fin_button['command']=self.mat_close
        if login_info['perm'][0]!='3' and login_info['perm'][0]!='9':
            self.widgets_group2.grid_forget()

        self.wl_str = StringVar()
        self.wl_entry = Entry(self, textvariable=self.wl_str)
        self.wl_entry.grid(row=0, column=6, sticky = EW)
        self.wl_entry.bind("<Return>", self.wl_result)
        model = TableModel(rows=0, columns=0)

        for colname in col_header:
            model.addColumn(colname)
        model.addRow(1)
        table_panel = Frame(self)
        self.mat_table = Table(table_panel, model, editable=False)
        self.mat_table.show()

        self.mat_table.grid(row=0, column=0, rowspan=2, columnspan=2, sticky=NSEW)
        table_panel.columnconfigure(1, weight=1)
        table_panel.rowconfigure(1, weight=1)
        table_panel.grid(row=1, column=0, rowspan=2, columnspan=8, sticky=NSEW )
        self.grid(row=0, column=0, sticky=NSEW)
        self.columnconfigure(7, weight=1)
        self.rowconfigure(2, weight=1)  
    
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
        model = TableModel(dataframe=self.df.T)
        self.mat_table.updateModel(model)
        self.mat_table.redraw()

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
            
        self.df = pd.DataFrame(self.mat_dic,index=col_header, columns=[ i for i in range(1, self.mr_rows+1)])
            
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
            
        self.df = pd.DataFrame(self.mat_dic,index=col_header, columns=[ i for i in range(1,self.mr_rows+1)])

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
    
        t_index= [ i for i in range(1, self.mr_rows+1)]
        self.df = pd.DataFrame(data=self.mat_dic,index=col_header, columns=t_index)
        model = TableModel(dataframe=self.df.T)
        self.mat_table.updateModel(model)
        self.mat_table.redraw()

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
        self.df = pd.DataFrame(self.mat_dic,index=col_header, columns=[ i for i in range(1, self.mr_rows+1)])
        model=TableModel(dataframe=self.df.T)
        self.mat_table.updateModel(model)
        self.mat_table.redraw()                     

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

#threads=[]
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
    data_thread = None
    b_finished=False
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
        self.refresh_button['command']=self.refresh_by_hand

        self.export_button=Button(self, text='导出EXCEL')
        self.export_button.grid(row=0, column=6,  sticky=EW)
        self.export_button['command']=self.__export_excel
        
        self.export_direct=Button(self, text='直接导出已完成清单\n(按CO run完成日期筛选)')
        self.export_direct.grid(row=0,column=7, sticky=EW)
        self.export_direct['command']=self.direct_export
        if login_info['perm'][1]!='1' and login_info['perm'][1]!='9':
            self.export_direct.grid_forget()        
              
        self.export_finished=Button(self, text='已完成清单\n(按CO run完成日期筛选)')
        self.export_finished.grid(row=1,column=7, sticky=EW)
        self.export_finished['command']=self.get_finished 
        if login_info['perm'][1]!='1' and login_info['perm'][1]!='9':
            self.export_finished.grid_forget()

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
        self.mat_list.grid(row=2, column=0, rowspan=2, columnspan=9, sticky='nsew')
        
        ysb.grid(row=2, column=9,rowspan=2, sticky='ns')
        xsb.grid(row=4, column=0, columnspan=9, sticky='ew')    
        self.grid()
        self.columnconfigure(8, weight=1)
        self.rowconfigure(3, weight=1)
       
        self.mat_list.bind('<Alt-c>', self.copy_mat_list)
        self.mat_list.bind('<Alt-C>', self.copy_mat_list)
        self.mat_list.bind('<Control-c>', self.copy_list)
        self.mat_list.bind('<Control-C>', self.copy_list) 
        
        self.focus_force()
        #self.bind('<Control-a>', self.select_all)
        #self.bind('<Control-A>', self.select_all)
        self.mat_list.bind('<Escape>', self.clear_select)
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
        if self.data_thread.is_alive():
            messagebox.showinfo('提示','表单刷新线程正在后台刷新列表，请等待完成后再点击!')
            return
            
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
        ws.cell(row=1, column=col_size+2).value = '合同号'
        ws.cell(row=1, column=col_size+3).value = '项目名称'
        ws.cell(row=1, column=col_size+4).value = '关联梯台数'
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
            if dic_str[nstd_id].startswith('E/'):
                wbs = self.nstd_res[nstd_id][0]
                prj_info = self.wbs_res[wbs]               
                ws.cell(row=n+2, column=col_size+2).value = prj_info[1]
                ws.cell(row=n+2, column=col_size+3).value = prj_info[2]
                
            ws.cell(row=n+2, column=col_size+4).value = len(self.nstd_res[nstd_id])
            n+=1

        if excel_xlsx.save_workbook(workbook=wb, filename=file_str):
            messagebox.showinfo("输出","成功输出!")  

    def direct_export(self):
        date_ctrl = date_picker(self)
        self.sel_range = date_ctrl.result
        if not self.sel_range:
            return
                   
        file_str=filedialog.asksaveasfilename(title="导出文件", filetypes=[('excel file','.xlsx')])
        if not file_str:
            return

        if not file_str.endswith(".xlsx"):
            file_str+=".xlsx"
            
        self.b_finished=True
        self.refresh_mat()  

        wb=Workbook()
        ws=wb.worksheets[0]
        ws.title='物料清单'
        col_size = len(cols)
        for i in range(col_size):
            ws.cell(row=1,column=i+1).value=tree_head[i]  
        ws.cell(row=1, column=col_size+1).value = '关联WBS'
        ws.cell(row=1, column=col_size+2).value = '合同号'
        ws.cell(row=1, column=col_size+3).value = '项目名称'
        n=0 
        
        dic_str = {}
        
        for key in range(len(self.mat_res)):
            for i in range(col_size):
                ws.cell(row=n+2, column=i+1).value = self.mat_res[key][i]

            j=cols.index('col1')
            nstd_id = self.mat_res[key][j]
            
            if nstd_id not in dic_str:
                s_str = self.combine_wbs(self.nstd_res[nstd_id])
                dic_str[nstd_id]=s_str
                
            ws.cell(row=n+2, column=col_size+1).value = dic_str[nstd_id]
            if dic_str[nstd_id].startswith('E/'):
                wbs = self.nstd_res[nstd_id][0]
                prj_info = self.wbs_res[wbs]               
                ws.cell(row=n+2, column=col_size+2).value = prj_info[1]
                ws.cell(row=n+2, column=col_size+3).value = prj_info[2]
            n+=1

        if excel_xlsx.save_workbook(workbook=wb, filename=file_str):
            messagebox.showinfo("输出","成功输出!") 
                 
    def get_finished(self):
        if self.data_thread.is_alive():
            messagebox.showinfo('提示','表单刷新线程正在后台刷新列表，请等待完成后再点击!')
            return
            
        date_ctrl = date_picker(self)
        self.sel_range = date_ctrl.result
        if not self.sel_range:
            return
            
        self.b_finished=True
        self.__refresh_tree()
    
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
            mat_items=nstd_app_head.select(nstd_app_head, nstd_mat_table, nstd_mat_fin).join(nstd_mat_table).switch(nstd_mat_table).join(nstd_mat_fin).where((nstd_mat_fin.justify==3)&(nstd_mat_fin.mbom_fin==False)&(nstd_mat_fin.co_run_fin==False)).naive()
        elif login_info['perm'][1]=='3':
            mat_items=nstd_app_head.select(nstd_app_head, nstd_mat_table, nstd_mat_fin).join(nstd_mat_table).switch(nstd_mat_table).join(nstd_mat_fin).where((nstd_mat_fin.justify==4)&(nstd_mat_fin.mbom_fin==False)&(nstd_mat_fin.co_run_fin==False)).naive()
        elif login_info['perm'][1]=='4':
            mat_items=nstd_app_head.select(nstd_app_head, nstd_mat_table, nstd_mat_fin).join(nstd_mat_table).switch(nstd_mat_table).join(nstd_mat_fin).where(((nstd_mat_fin.justify==1)|(nstd_mat_fin.justify==2)|(nstd_mat_fin.justify==6))&(nstd_mat_fin.pu_price_fin==False)&(nstd_mat_fin.co_run_fin==False)).naive()
        elif login_info['perm'][1]=='5':
            mat_items=nstd_app_head.select(nstd_app_head, nstd_mat_table, nstd_mat_fin).join(nstd_mat_table).switch(nstd_mat_table).join(nstd_mat_fin)\
                .where(((((nstd_mat_fin.justify==1)|(nstd_mat_fin.justify==2)|(nstd_mat_fin.justify==6))&(nstd_mat_fin.pu_price_fin==True))|(((nstd_mat_fin.justify==3)|(nstd_mat_fin.justify==4)|(nstd_mat_fin.justify==5))&(nstd_mat_fin.mbom_fin==True)))&(nstd_mat_fin.co_run_fin==False)).naive()
        elif login_info['perm'][1]=='6':
            mat_items=nstd_app_head.select(nstd_app_head, nstd_mat_table, nstd_mat_fin).join(nstd_mat_table).switch(nstd_mat_table).join(nstd_mat_fin).where((nstd_mat_fin.justify==5)&(nstd_mat_fin.mbom_fin==False)&(nstd_mat_fin.co_run_fin==False)).naive()  
        elif login_info['perm'][1]=='1' or login_info['perm'][1]=='9':
            if not self.b_finished:
                mat_items=nstd_app_head.select(nstd_app_head, nstd_mat_table, nstd_mat_fin).join(nstd_mat_table).switch(nstd_mat_table).join(nstd_mat_fin).where((nstd_mat_fin.justify>=0)&(nstd_mat_fin.co_run_fin==False)).naive()
            else:
                mat_items=nstd_app_head.select(nstd_app_head, nstd_mat_table, nstd_mat_fin).join(nstd_mat_table).switch(nstd_mat_table).join(nstd_mat_fin).where((nstd_mat_fin.justify>=0)&(nstd_mat_fin.co_run_fin==True)&(nstd_mat_fin.co_run_fin_on>=self.sel_range['from'])&(nstd_mat_fin.co_run_fin_on<=self.sel_range['to'])).naive()
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
            result = LProcAct.select(LProcAct.finish_date).where((LProcAct.instance==index)&(LProcAct.action=='AT00000015')).get()
            s_date = result.finish_date
        except LProcAct.DoesNotExist:
            s_date=''

        item.append(s_date)

        try:
            result = LProcAct.select(LProcAct.finish_date).where((LProcAct.instance==index)&(LProcAct.action=='AT00000020')).get()
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
        self.data_thread= refresh_thread(self)
        self.data_thread.setDaemon(True)
        self.data_thread.start()
        #threads.append(data_thread)
    
    def refresh_by_hand(self):
        if self.data_thread.is_alive():
            messagebox.showinfo('提示','表单刷新线程正在后台刷新列表，请等待完成后再点击!')
            return
            
        self.b_finished=False
        self.__refresh_tree()

    def __loop_refresh(self):
        if not self.data_thread:
            pass
        elif self.data_thread.is_alive():
            messagebox.showinfo('提示','表单刷新线程正在后台运行, 15分钟后自动刷新进程重新启动!')
            time.sleep(900)
            
        self.b_finished=False
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
    '''
    def copy_clip_process(self):
        wbs_sel = WBSSelDialog(self, 'WBS List粘贴')
    '''        
    def __mat_list(self,choice, mats, comment=None):
        s_error=''
        i_error=0
        i_suss=0
        for mat in mats:
            result = self.process_mat_status(mat.rstrip(), choice, True)
            if result >0:
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

        count = len(items)
        if messagebox.askyesno('确认执行','执行数据数量: '+str(count)+' 条;此操作不可逆，是否继续(YES/NO)?')==NO:
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

    def process_mat_status(self, mat, method, value=True, comment=None):
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
        except nstd_mat_fin.DoesNotExist:
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
                    s = nstd_mat_fin.update(mbom_fin=value,mbom_fin_on=datetime.datetime.now(), mbom_fin_by=login_info['uid'], mbom_fin_remark=none2str(comment)).where(nstd_mat_fin.mat_no==mat)
                    i_u=s.execute()               
            else:
                return -2
        elif method==2:
            if (mat_catalog==1 or mat_catalog==2 or mat_catalog==6) and (user_per==4 or user_per==9):
                if mat_pu_price_fin ==value:
                    return 2
                else:
                    s = nstd_mat_fin.update(pu_price_fin=value,pu_price_fin_on=datetime.datetime.now(), pu_price_fin_by=login_info['uid'], pu_price_fin_remark=none2str(comment)).where(nstd_mat_fin.mat_no==mat)
                    i_u=s.execute()  
            else: 
                return -2
        elif method==3:
            if user_per!=5 and user_per!=9:
                return -2
            
            if mat_catalog<=0:
                return -3

            if (mat_catalog==1 or mat_catalog==2 or mat_catalog==6) and not mat_pu_price_fin:
                return -3

            if (mat_catalog==3 or mat_catalog==4 or mat_catalog==5) and not mat_mbom_fin:
                return -3

            if mat_co_run_fin == value:
                return 2
            else:
                s = nstd_mat_fin.update(co_run_fin=value,co_run_fin_on=datetime.datetime.now(), co_run_fin_by=login_info['uid'], co_run_fin_remark=none2str(comment)).where(nstd_mat_fin.mat_no==mat)
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
    prj_thread = None
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
        self.refresh_button['command']=self.refresh_by_hand
        
        self.release_button=Button(self, text='项目release')
        self.release_button.grid(row=0, column=2, sticky=NSEW)
        self.release_button['command']= self.release_prj
        if login_info['perm'][2]!='2' and login_info['perm'][2]!='9':
            self.release_button.grid_forget() 
            
        self.export_button=Button(self, text='导出清单')
        self.export_button.grid(row=0, column=3, sticky=NSEW)
        self.export_button['command']=self.proj_export
        
        self.delivery_button=Button(self, text='发运完成')
        self.delivery_button.grid(row=1, column=3, sticky=NSEW)
        self.delivery_button['command']=self.delivery
        if login_info['perm'][2]!='3' and login_info['perm'][2]!='9':
            self.delivery_button.grid_forget()  
            
        self.restore_prj=Button(self, text='恢复项目到待排产状态')
        self.restore_prj.grid(row=0, column=4, sticky=NSEW)
        self.restore_prj['command']=self.restore_prj_proc
        if login_info['perm'][2]!='3' and login_info['perm'][2]!='9':
            self.restore_prj.grid_forget()
            
        self.update_date_button = Button(self, text='更新非标交货期和配置完成日期')
        self.update_date_button.grid(row=1, column=4, sticky=NSEW)
        self.update_date_button['command']=self.update_date
        if login_info['perm'][2]!='3' and login_info['perm'][2]!='9':
            self.update_date_button.grid_forget()
            
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
        
        #self.bind('<Control-a>', self.select_all)
        #self.bind('<Control-A>', self.select_all)
        self.wbs_list.bind('<Escape>', self.clear_select)
        self.wbs_list.bind('<Control-a>', self.select_all)
        self.wbs_list.bind('<Control-A>', self.select_all)
        
        if login_info['status'] and int(login_info['perm'][2])>=1:
            self.__loop_refresh()
            
    def select_all(self, event):
        items = self.wbs_list.get_children()
        if not items:
            return

        self.wbs_list.selection_set(items[0])   
        self.wbs_list.focus_set()
        self.wbs_list.focus(items[0])
        for item in items:
            self.wbs_list.selection_add(item)
    
    def clear_select(self, event):
        items=self.wbs_list.selection()
        for item in items:
            self.wbs_list.selection_remove(item)
            
    def proj_export(self):
        if self.prj_thread.is_alive():
            messagebox.showinfo('提示','表单刷新线程正在后台刷新列表，请等待完成后再点击!')
            return
        
        items = self.wbs_list.selection()
        if not items:        
            return
 
        file_str=filedialog.asksaveasfilename(title="导出文件", filetypes=[('excel file','.xlsx')])
        if not file_str:
            return

        if not file_str.endswith(".xlsx"):
            file_str+=".xlsx"

        wb=Workbook()
        ws0=wb.worksheets[0]
        
        ws0.title='项目清单'
        col_size = len(project_head)
        for i in range(col_size):
            ws0.cell(row=1,column=i+1).value=project_head[i]
        
        rid=0
        wbses=[]
        for item in items:
            value = self.wbs_list.item(item, 'values')
            wbses.append(value[1])
            cid=0
            for c_str in value:
                ws0.cell(row=rid+2,column=cid+1).value=value[cid]
                cid+=1
            rid+=1
        
        col_header=['导入日期','非标编号','判断','物料号','物料名称(中)','物料名称(英)','图号','单位','备注','RP','BoxId','申请人','自制完成','负责人','完成日期','价格维护','负责人','完成日期','CO-run','负责人','完成日期']
        ws1=wb.create_sheet()        
        ws1.title='PSM负责'
        col_size=len(col_header)
        
        for col in range(col_size):
            ws1.cell(row=1, column=col+1).value = col_header[col]
        
        mat_items= nstd_app_head.select(nstd_app_head, nstd_mat_table, nstd_mat_fin).join(nstd_app_link).switch(nstd_app_head).join(nstd_mat_table).switch(nstd_mat_table).join(nstd_mat_fin).\
                        where((nstd_app_link.wbs_no.in_(wbses))&((nstd_mat_fin.justify==1)|(nstd_mat_fin.justify==2)|(nstd_mat_fin.justify==6))&(nstd_mat_fin.pu_price_fin==False)&(nstd_mat_fin.co_run_fin==False)).naive()

        self.create_excel_sheet(ws1, mat_items)
        
        ws2= wb.create_sheet()
        ws2.title='钣金自制'
        for col in range(col_size):
            ws2.cell(row=1, column=col+1).value = col_header[col]
        
        mat_items= nstd_app_head.select(nstd_app_head, nstd_mat_table, nstd_mat_fin).join(nstd_app_link).switch(nstd_app_head).join(nstd_mat_table).switch(nstd_mat_table).join(nstd_mat_fin).\
                        where((nstd_app_link.wbs_no.in_(wbses))&(nstd_mat_fin.justify==4)&(nstd_mat_fin.mbom_fin==False)&(nstd_mat_fin.co_run_fin==False)).naive()

        self.create_excel_sheet(ws2, mat_items)
        
        ws3 =wb.create_sheet()
        ws3.title = '电气自制'
        for col in range(col_size):
            ws3.cell(row=1, column=col+1).value = col_header[col]
        
        mat_items= nstd_app_head.select(nstd_app_head, nstd_mat_table, nstd_mat_fin).join(nstd_app_link).switch(nstd_app_head).join(nstd_mat_table).switch(nstd_mat_table).join(nstd_mat_fin).\
                        where((nstd_app_link.wbs_no.in_(wbses))&(nstd_mat_fin.justify==3)&(nstd_mat_fin.mbom_fin==False)&(nstd_mat_fin.co_run_fin==False)).naive()

        self.create_excel_sheet(ws3, mat_items)      

        ws4 = wb.create_sheet()
        ws4.title = '曳引机自制'
        for col in range(col_size):
            ws4.cell(row=1, column=col+1).value = col_header[col]
        
        mat_items= nstd_app_head.select(nstd_app_head, nstd_mat_table, nstd_mat_fin).join(nstd_app_link).switch(nstd_app_head).join(nstd_mat_table).switch(nstd_mat_table).join(nstd_mat_fin).\
                        where((nstd_app_link.wbs_no.in_(wbses))&(nstd_mat_fin.justify==5)&(nstd_mat_fin.mbom_fin==False)&(nstd_mat_fin.co_run_fin==False)).naive()

        self.create_excel_sheet(ws4, mat_items) 
        
        ws5 =wb.create_sheet()
        ws5.title = 'CO 负责'
        for col in range(col_size):
            ws5.cell(row=1, column=col+1).value = col_header[col]
        
        mat_items= nstd_app_head.select(nstd_app_head, nstd_mat_table, nstd_mat_fin).join(nstd_app_link).switch(nstd_app_head).join(nstd_mat_table).switch(nstd_mat_table).join(nstd_mat_fin).\
                        where((nstd_app_link.wbs_no.in_(wbses))&(nstd_mat_fin.justify>0)&(nstd_mat_fin.co_run_fin==False)).naive()
        
        self.create_excel_sheet(ws5, mat_items) 
        
        if excel_xlsx.save_workbook(workbook=wb, filename=file_str):
            messagebox.showinfo("输出","成功输出!")  
        
            
    def create_excel_sheet(self, ws,  mat_items):
        rid=0
        for r in mat_items:
            ws.cell(row=rid+2, column=1).value=none2str(r.modify_on)
            ws.cell(row=rid+2, column=2).value=r.nstd_app
            ws.cell(row=rid+2, column=3).value=Justify_Types[r.justify]
            ws.cell(row=rid+2, column=4).value= r.mat_no
            ws.cell(row=rid+2, column=5).value = r.mat_name_cn
            ws.cell(row=rid+2, column=6).value = r.mat_name_en
            ws.cell(row=rid+2, column=7).value = none2str(r.drawing_no)
            ws.cell(row=rid+2, column=8).value = r.mat_unit
            ws.cell(row=rid+2, column=9).value = none2str(r.comments)
            ws.cell(row=rid+2, column=10).value = none2str(r.rp)
            ws.cell(row=rid+2, column=11).value = none2str(r.box_code_sj)
            app_per = r.app_person
            if app_per is None:
                app_per =none2str(r.mat_app_person)
            elif len(app_per)==0:
                app_per =none2str(r.mat_app_person)
                
            ws.cell(row=rid+2, column=12).value = app_per
            
            temp = (r.mbom_fin and 'Y' or '')
            ws.cell(row=rid+2, column=13).value = temp
            if temp.upper()=='Y':
                ws.cell(row=rid+2, column=14).value = get_name(r.mbom_fin_by)
                ws.cell(row=rid+2, column=15).value = none2str(r.mbom_fin_on)
            else:
                ws.cell(row=rid+2, column=14).value =''
                ws.cell(row=rid+2, column=15).value = ''
                
            temp=(r.pu_price_fin and 'Y' or '')
            ws.cell(row=rid+2, column=16).value = temp
            if temp.upper()=='Y':
                ws.cell(row=rid+2, column=17).value = get_name(r.pu_price_fin_by)
                ws.cell(row=rid+2, column=18).value = none2str(r.pu_price_fin_on)
            else:
                ws.cell(row=rid+2, column=17).value = ''
                ws.cell(row=rid+2, column=18).value = ''
                
            temp = (r.co_run_fin and 'Y' or '')
            ws.cell(row=rid+2, column=19).value = temp
            if temp.upper()=='Y':
                ws.cell(row=rid+2, column=20).value = get_name(r.co_run_fin_by)
                ws.cell(row=rid+2, column=21).value = none2str(r.co_run_fin_on)
            else:
                ws.cell(row=rid+2, column=20).value = ''
                ws.cell(row=rid+2, column=21).value = ''
                
            rid+=1
            
    
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
    
    def update_date(self):
        d = ask_list('物料拷贝器', 1)
        if not d:
            return
        
        for item in d:
            self.update_delivery_date(item)
            
    def update_delivery_date(self, wbs):
        try:
            r = UnitInfo.get(UnitInfo.wbs_no==wbs)
        except UnitInfo.DoesNotExist:   
            return
        
        configure_date = r.req_configure_finish
        
        q_res=  NonstdAppHeader.update(drawing_req_date=configure_date).where((NonstdAppHeader.link_list.contains(wbs))&(NonstdAppHeader.status>=0))
        q_res.execute()
            
    def do_popup(self, event):
        # display the popup menu
        try:
            self.popup.tk_popup(event.x_root, event.y_root, 0)
        finally:
            # make sure to release the grab (Tk 8.0a1 only)
            self.popup.grab_release() 
            
    def restore_prj_proc(self):
        d = ask_list('物料拷贝器', 1)
        if not d:
            return
        
        for item in d:
            self.restore_unit(item)
    
    def restore_unit(self,wbs):
        q_res = NonstdAppItem.select(NonstdAppItem.index,NonstdAppItem.nonstd,NonstdAppItemInstance.index_mat,NonstdAppItemInstance.has_nonstd_draw,NonstdAppItemInstance.has_nonstd_mat).join(NonstdAppItemInstance, JOIN.LEFT_OUTER, on=(NonstdAppItem.index==NonstdAppItemInstance.index)).\
                   where(NonstdAppItem.link_list.contains(wbs)).naive()
        
        if not q_res:
            pass
        else:
            for r in q_res:
                if r.has_nonstd_draw:
                    q=LProcAct.delete().where((LProcAct.instance==r.index_mat)&(LProcAct.workflow=='WF0004'))
                    q.execute()
                
                if r.has_nonstd_mat:
                    q=LProcAct.delete().where((LProcAct.instance==r.index_mat)&(LProcAct.workflow=='WF0005'))
                    q.execute()
                    
                q=LProcAct.delete().where((LProcAct.instance==r.index)&(LProcAct.workflow=='WF0003'))
                q.execute()
                                   
                q = NonstdAppItemInstance.update(status=-1).where(NonstdAppItemInstance.index_mat==r.index_mat)
                q.execute()
                
                q= NonstdAppItem.update(status=-1).where(NonstdAppItem.index==r.index)
                q.execute()
                
                q= NonstdAppHeader.update(status=-1).where(NonstdAppHeader.nonstd==r.nonstd)
                q.execute()
                
                
        q=LProcAct.update(is_active=False).where((LProcAct.instance==wbs)&(LProcAct.is_active==True)&((LProcAct.workflow=='WF0002')|(LProcAct.workflow=='WF0006')))
        q.execute()
        
        q = UnitInfo.update(status=2).where(UnitInfo.wbs_no==wbs)  
        q.execute()     
                       
    def delivery(self):
        sel_items = self.wbs_list.selection()
        if not sel_items: 
            d = ask_list('物料拷贝器', 1)
            if not d:
                return
            self.make_delivery(d)            
            return

        if self.sel_index !=1:
            messagebox.showwarning('提示', '请在选项中选择已经release的项目')
            return       
         
        error=[]
        for item in sel_items:
            i = project_cols.index('col2')
            wbs = self.wbs_list.item(item, 'values')[i]
            
            q = UnitInfo.update(status=8, modify_emp=login_info['uid'], modify_date=datetime.datetime.now()).where(UnitInfo.wbs_no==wbs)
            j= q.execute() 
            if j==0:
                error.append(wbs)
            else:
                self.wbs_dic.pop(wbs)
                self.wbs_list.delete(item)
                self.wbs_keys.remove(wbs)
        i_count=len(sel_items)
        i_fail =len(error)
        
        messagebox.showinfo('delivery结果','delivery成功:'+str(i_count-i_fail)+';失败:'+str(i_fail))
        
    def make_delivery(self, wbses):
        error=[]
        e_str=''
        for wbs in wbses:
            try:
                UnitInfo.get((UnitInfo.wbs_no==wbs)&(UnitInfo.status==6))
            except UnitInfo.DoesNotExist:
                error.append(wbs)
                e_str=e_str+'\n'+wbs+'状态未RELEASE'
                continue
           
            q = UnitInfo.update(status=8, modify_emp=login_info['uid'], modify_date=datetime.datetime.now()).where(UnitInfo.wbs_no==wbs)
            j= q.execute() 
            if j==0:
                error.append(wbs)
                e_str=e_str+'\n'+wbs+'更新失败'
        
        i_count=len(wbses)
        i_fail =len(error)
                
        messagebox.showinfo('delivery结果','delivery成功:'+str(i_count-i_fail)+';失败:'+str(i_fail)+':'+e_str)            
                
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
        if self.prj_thread.is_alive():
            messagebox.showinfo('提示','表单刷新线程正在后台刷新列表，请等待完成后再点击!')
            return
                
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
        self.prj_thread= refresh_thread(self)
        self.prj_thread.setDaemon(True)
        self.prj_thread.start()
        #threads.append(prj_thread)
        
    def refresh_by_hand(self):
        if self.prj_thread.is_alive():
            messagebox.showinfo('提示','表单刷新线程正在后台刷新列表，请等待完成后再点击!')
            return
            
        self.__refresh_list()

    def __loop_refresh(self):
        if not self.prj_thread:
            pass
        elif self.prj_thread.is_alive():
            messagebox.showinfo('提示','表单刷新线程正在后台运行, 15分钟后自动刷新进程重新启动!')
            time.sleep(900)
            
        self.__refresh_list()
        self.wbs_list.after(3600000, self.__loop_refresh)
        
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
            self.proc_list.heading(col, text=self.tree_head_his[i], command=lambda _col=col: treeview_sort_column(self.proc_list, _col, False))  
        
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
        


logger = logging.getLogger()
 
mat_heads = ['位置号','物料号','中文名称','英文名称','图号','数量','单位','材料','重量','备注']
mat_keys = ['st_no','mat_no','mat_name_cn','mat_name_en','drawing_no','qty','mat_unit','mat_material','part_weight','comments']

mat_cols = ['col1','col2','col3','col4','col5','col6','col7','col8','col9','col10']

def tree_level(val):
    l = len(val)
    if l==0:
        return 0
    
    r=1
    for i in range(l):
        if int(val[i])>0:
            return r
        elif int(val[i])==0:
            r+=1
            
    return r

def dict2list(dict):
    li = []
    
    for i in range(len(mat_heads)):
        li.append(dict[mat_heads[i]])
    
    return li

def cell2str(val):
    if (val is None) or (val=='N/A') or (val=='N') or (val=='无'):
        return ''
    else:
        return str(val).strip()
    
class TextHandler(logging.Handler):
    """This class allows you to log to a Tkinter Text or ScrolledText widget"""
    def __init__(self, text):
        # run the regular Handler __init__
        logging.Handler.__init__(self)
        # Store a reference to the Text it will log to
        self.text = text

    def emit(self, record):
        self.formatter = logging.Formatter('%(asctime)s-%(levelname)s : %(message)s')
        msg = self.format(record)
        def append():
            self.text.configure(state='normal')          
            self.text.insert(END, msg+"\n")
            self.text.configure(state='disabled')
            # Autoscroll to the bottom
            self.text.yview(END)
        # This is necessary because we can't modify the Text from other threads
        self.text.after(0, append)# Scroll to the bottom

class eds_pane(Frame):
    '''
    mat_list = {1:{'位置号':value,'物料号':value, ....,'标判断':value},.....,item:{......}}
    bom_tree : 物料BOM的树形结构, 以key为节点,保存om树型结构如下:
    0
    ├── 1
    │   └── 3
    └── 2
    '''
    hibe_mats=[]
    mat_list = {} #从文件读取的文件列表，以 数字1，2，...为keys
    bom_items = [] #存储有下层BOM的节点，treeview 控件的节点
    mat_items = {} #以物料号为key,存储涉及BOM的物料清单 ，包括下层物料。
    #treeview本身是树形结构，无需在重新构建树形model
    #bom_tree = Tree()
    #bom_tree.create_node(0,0)
    mat_pos = 0 #配合mat_list的的数量
    mat_tops={} #发运层物料字典，key为物料号，value是struct code 和revision列表
    nstd_mat_list=[] #非标物料列表
    sap_thread = None
    nstd_app_id=''
    def __init__(self,master=None):
        Frame.__init__(self, master)
        self.grid()
        
        self.createWidgets()
        
    def createWidgets(self):
        '''
        self.find_mode = StringVar()
        self.find_combo = ttk.Combobox(self,textvariable = self.find_mode)
        self.find_combo['values'] = ('列出物料BOM结构','查找物料的上层','查找物料关联项目','查找项目关联物料')
        self.find_combo.current(0)
        self.find_combo.grid(row =0,column=0, columnspan=2,sticky=EW)
        '''
        self.find_label = Label(self, text='请输入头物料号查找',anchor='w') 
        self.find_label.grid(row=0, column=0, columnspan=2, sticky=EW)
              
        self.find_var = StringVar()
        self.find_text = Entry(self, textvariable=self.find_var)
        self.find_text.grid(row=1, column=0, columnspan=2, sticky=EW)
        self.find_text.bind("<Return>", self.search)
        
        self.version_label = Label(self,text='物料版本', anchor ='w')
        self.version_label.grid(row=0, column=2, columnspan=2, sticky=EW)
        
        self.version_var = StringVar()
        self.version_text = Entry(self, textvariable=self.version_var)
        self.version_text.grid(row=1, column=2, columnspan=2, sticky=EW)
        self.version_text.bind("<Return>", self.search)
                
        self.st_body = Frame(self)
        self.st_body.grid(row=0, column=4, rowspan=2, columnspan=4, sticky=NSEW)
        
        self.import_button = Button(self.st_body, text='文件读取')
        self.import_button.grid(row=0, column=0, sticky=NSEW)
        self.import_button['command'] = self.excel_import
        
        self.generate_nstd_list = Button(self.st_body, text='生成非标物料申请表')
        self.generate_nstd_list.grid(row=1, column=0, sticky=NSEW)
        self.generate_nstd_list['command']=self.generate_app
        
        self.pdm_generate_button = Button(self.st_body, text='PDM物料导入文件生成\n(物料清单和BOM清单)')
        self.pdm_generate_button.grid(row=0, column=1, rowspan=2, sticky=NSEW)
        self.pdm_generate_button['command'] = self.pdm_generate
        
        self.para_label = Label(self.st_body,text='搜索参数', anchor ='w')
        self.para_label.grid(row=0, column=2, columnspan=2, sticky=EW)
        
        self.para_var = StringVar()
        self.para_text = Entry(self.st_body, textvariable=self.para_var)
        self.para_text.grid(row=1, column=2, columnspan=2, sticky=EW)
        self.para_text.bind("<Return>", self.para_search)        
               
        self.ie_body = Frame(self)
        self.ie_body.grid(row=0, column=8,rowspan=2, columnspan=4, sticky=NSEW)
        
        self.multi_search = Button(self.ie_body, text='多物料搜索')
        self.multi_search.grid(row=0, column=0, sticky=NSEW)
        self.multi_search['command']= self.multi_find
        
        self.import_bom_List = Button(self.ie_body, text='生成BOM导入表')
        self.import_bom_List.grid(row=1, column=0, sticky=NSEW)
        self.import_bom_List['command']=self.import_bom_list_x
        
        self.ntbook = ttk.Notebook(self)        
        self.ntbook.rowconfigure(0, weight=1)
        self.ntbook.columnconfigure(0, weight=1)
        '''
                清单式显示不够直观，同时pandastable表操作速度太慢，故只使用树形结构
        list_pane = Frame(self)
        model = TableModel(rows=0, columns=0)
        for col in mat_heads:
            model.addColumn(col)
        model.addRow(1)
        
        self.mat_table = Table(list_pane, model, editable=False)
        self.mat_table.show()
        '''                
        tree_pane = Frame(self)
        self.mat_tree = ttk.Treeview(tree_pane, columns=mat_cols,selectmode='extended')
        style = ttk.Style()
        style.configure("Treeview", font=('TkDefaultFont', 10))
        style.configure("Treeview.Heading", font=('TkDefaultFont', 9))  
        self.mat_tree.heading('#0', text='')
        for col in mat_cols:
            i = mat_cols.index(col)
            if i==0:
                self.mat_tree.heading(col,text="版本号/位置号")
            else:
                self.mat_tree.heading(col,text=mat_heads[i])
        
        #('位置号','物料号','中文名称','英文名称','图号','数量','单位','材料','重量','备注')
        self.mat_tree.column('#0', width=80)
        self.mat_tree.column('col1', width=80, anchor='w')
        self.mat_tree.column('col2', width=100, anchor='w')
        self.mat_tree.column('col3', width=150, anchor='w')
        self.mat_tree.column('col4', width=150, anchor='w')
        self.mat_tree.column('col5', width=100, anchor='w')
        self.mat_tree.column('col6', width=100, anchor='w')
        self.mat_tree.column('col7', width=100, anchor='w')
        self.mat_tree.column('col8', width=150, anchor='w')
        self.mat_tree.column('col9', width=100, anchor='w')
        self.mat_tree.column('col10', width=300, anchor='w')  
               
        ysb = ttk.Scrollbar(tree_pane, orient='vertical', command=self.mat_tree.yview)
        xsb = ttk.Scrollbar(tree_pane, orient='horizontal', command=self.mat_tree.xview)        
        ysb.grid(row=0, column=2, rowspan=2, sticky='ns')
        xsb.grid(row=2, column=0, columnspan=2, sticky='ew')  
        
        self.mat_tree.configure(yscroll=ysb.set, xscroll=xsb.set)
        self.mat_tree.grid(row=0, column=0, rowspan=2, columnspan=2, sticky='nsew')
        tree_pane.rowconfigure(1, weight=1)
        tree_pane.columnconfigure(1, weight =1)
        
        #self.ntbook.add(list_pane, text='BOM清单', sticky=NSEW)
        self.ntbook.add(tree_pane, text='BOM树形结构', sticky=NSEW) 
        
        log_pane = Frame(self)
        
        self.log_label=Label(log_pane)
        self.log_label["text"]="操作记录"
        self.log_label.grid(row=0,column=0, sticky=W)
        
        self.log_text=scrolledtext.ScrolledText(log_pane, state='disabled')
        self.log_text.config(font=('TkFixedFont', 10, 'normal'))
        self.log_text.grid(row=1, column=0, columnspan=2, sticky=EW)
        log_pane.rowconfigure(1,weight=1)
        log_pane.columnconfigure(1, weight=1)
        
        self.ntbook.grid(row=2, column=0, rowspan=6, columnspan=12, sticky=NSEW)
        log_pane.grid(row=8, column=0,columnspan=12, sticky=NSEW)
              
        # Create textLogger
        text_handler = TextHandler(self.log_text)        
        # Add the handler to logger
        
        logger.addHandler(text_handler)
        logger.setLevel(logging.INFO) 
        
        self.rowconfigure(8, weight=1)
        self.columnconfigure(11, weight=1) 
        
        if login_info['perm'][3]!='1' and login_info['perm'][3]!='9':
            self.st_body.grid_forget()
            
        if login_info['perm'][3]!='2' and login_info['perm'][3]!='9':
            self.ie_body.grid_forget()
    
    def pdm_generate(self):
        if len(self.bom_items)==0:
            logger.warning('没有bom结构，请先搜索物料BOM')
            return
    
        if self.sap_thread is not None and self.sap_thread.is_alive():
            messagebox.showinfo('提示','正在后台检查SAP非标物料，请等待完成后再点击!')
            return  
        
        if len(self.nstd_mat_list) == 0:
            logger.warning('此物料BOM中包含未维护进SAP系统的物料，请等待其维护完成')
            return
        
        if len(self.nstd_app_id)==0:
            logger.warning('请先生成非标申请表，填入非标单号后生成此文件')
            return
        
        gen_dir = filedialog.askdirectory(title="请选择输出文件保存的文件夹!")
        if not gen_dir or len(gen_dir)==0:
            return        

        temp_file = os.path.join(cur_dir(),'PDMT1.xls')
        rb = xlrd.open_workbook(temp_file, formatting_info=True)
                
        wb= copy(rb)
        ws = wb.get_sheet(0)
        
        #now = datetime.datetime.now()
        #s_now = now.strftime('%Y%m%d%H%M%S')
        file_name = self.nstd_app_id+'物料清单.xls'
        pdm_mats_str =os.path.join(gen_dir, file_name)
        logger.info('正在生成导入物料清单文件:'+pdm_mats_str)
        i=2
        for it in self.nstd_mat_list:
            ws.write(i,0, it)
            value = self.mat_items[it]
            
            ws.write(i,1, value[mat_heads[4]])
            ws.write(i,3, value[mat_heads[2]])
            ws.write(i,4, value[mat_heads[3]])
            ws.write(i, 6, value[mat_heads[7]])
            ws.write(i,8, value[mat_heads[6]])
            ws.write(i, 9, value[mat_heads[9]])
            ws.write(i, 11, 'EDS系统')
            
            if it in self.mat_tops:
                rp_box = self.mat_tops[it]['rp_box']
                ws.write(i,12,rp_box['2101'][0])
                ws.write(i,13,rp_box['2101'][1])
                ws.write(i,14,rp_box['2001'][0])
                ws.write(i,15,rp_box['2001'][1])
 
            i+=1
                        
        wb.save(pdm_mats_str)
        logger.info(pdm_mats_str+'保存完成!')
        
        temp_file = os.path.join(cur_dir(),'PDMT2.xlsx')
        logger.info('正在根据模板文件:'+temp_file+'生成PDM BOM导入清单...')
        wb = load_workbook(temp_file) 
        temp_ws = wb.get_sheet_by_name('template')
        
        for it in self.bom_items:
            p_mat = self.mat_tree.item(it, 'values')[1]
            ws = wb.copy_worksheet(temp_ws)     
            ws.sheet_state ='visible'
            ws.title = p_mat
            
            logger.info('正在构建物料'+p_mat+'的PDM BOM导入清单...')
            p_name = self.mat_tree.item(it, 'values')[2]
            p_drawing = self.mat_tree.item(it,'values')[4]
            ws.cell(row=43, column=18).value = p_mat
            ws.cell(row=41, column=18).value = p_name
            ws.cell(row=45, column=18).value = p_drawing
            ws.cell(row=41, column=10).value = 'L'+p_mat
            children = self.mat_tree.get_children(it)
            i=4
            for child in children:
                value = self.mat_tree.item(child, 'values')
                ws.cell(row=i, column=2).value = value[1]
                ws.cell(row=i, column=5).value = value[4]
                ws.cell(row=i, column=10).value = value[6]
                ws.cell(row=i, column=13).value = value[2]
                ws.cell(row=i, column=16).value = value[7]
                ws.cell(row=i, column=20).value = value[5]
                ws.cell(row=i, column=23).value = value[9]
                i+=1
                
        wb.remove_sheet(temp_ws)
        file_name = self.nstd_app_id+'PDM BOM物料导入清单.xlsx'
        pdm_bom_str = os.path.join(gen_dir, file_name)
        if writer.excel.save_workbook(workbook=wb, filename=pdm_bom_str):
            logger.info('生成PDM BOM导入清单:'+pdm_bom_str+' 成功!')
        else:
            logger.info('文件保存失败!')
    
    def import_bom_list_x(self): 
        if len(self.bom_items)==0:
            logger.warning('没有bom结构，请先搜索物料BOM')
            return
    
        if self.sap_thread is not None and self.sap_thread.is_alive():
            messagebox.showinfo('提示','正在后台检查SAP非标物料，请等待完成后再点击!')
            return  
        
        if len(self.nstd_mat_list) != 0:
            logger.warning('此物料BOM中包含未维护进SAP系统的物料，请等待其维护完成')
            return
        
        file_str=filedialog.asksaveasfilename(title="导出文件", initialfile="temp",filetypes=[('excel file','.xlsx')])
        if not file_str:
            return

        if not file_str.endswith(".xlsx"):
            file_str+=".xlsx"        

        temp_file = os.path.join(cur_dir(),'bom.xlsx')
        wb = load_workbook(temp_file) 
        ws = wb.get_sheet_by_name('BOM')
        
        logger.info('正在生成文件'+file_str)
        i=4
        for it in self.bom_items:
            p_mat = self.mat_tree.item(it, 'values')[1]
            logger.info('正在构建物料'+p_mat+'的BOM导入清单...')
            p_name = self.mat_tree.item(it, 'values')[2]
            children = self.mat_tree.get_children(it)
            for child in children:
                value = self.mat_tree.item(child, 'values')
                c_mat = value[1]
                c_name = value[2]
                ws.cell(row=i, column=1).value= p_mat
                ws.cell(row=i, column=2).value= p_name
                ws.cell(row=i, column=6).value= c_mat
                ws.cell(row=i, column=7).value= c_name
                ws.cell(row=i, column=3).value= 2102
                ws.cell(row=i, column=4).value= 1
                if c_mat in self.hibe_mats:
                    ws.cell(row=i, column=5).value= 'N'
                else:
                    ws.cell(row=i, column=5).value= 'L'
                    ws.cell(row=i, column=15).value = 'X'
                    
                ws.cell(row=i, column=8).value=float(value[5])
                i+=1
                        
        if writer.excel.save_workbook(workbook=wb, filename=file_str):
            logger.info('生成BOM导入清单文件:'+file_str+' 成功!')
        else:
            logger.info('文件保存失败!')
                     
    def import_bom_list(self):
        if len(self.bom_items)==0:
            logger.warning('没有bom结构，请先搜索物料BOM')
            return
    
        if self.sap_thread is not None and self.sap_thread.is_alive():
            messagebox.showinfo('提示','正在后台检查SAP非标物料，请等待完成后再点击!')
            return  
        
        if len(self.nstd_mat_list) != 0:
            logger.warning('此物料BOM中包含未维护进SAP系统的物料，请等待其维护完成')
            return
        
        file_str=filedialog.asksaveasfilename(title="导出文件", initialfile="temp",filetypes=[('excel file','.xls')])
        if not file_str:
            return

        if not file_str.endswith(".xls"):
            file_str+=".xls"        

        temp_file = os.path.join(cur_dir(),'bom.xls')
        rb = xlrd.open_workbook(temp_file, formatting_info=True)
                
        wb= copy(rb)
        ws = wb.get_sheet(0)
        
        logger.info('正在生成文件'+file_str)
        i=4
        for it in self.bom_items:
            p_mat = self.mat_tree.item(it, 'values')[1]
            logger.info('正在构建物料'+p_mat+'的BOM导入清单...')
            p_name = self.mat_tree.item(it, 'values')[2]
            children = self.mat_tree.get_children(it)
            for child in children:
                value = self.mat_tree.item(child, 'values')
                c_mat = value[1]
                c_name = value[2]
                ws.write(i, 0, p_mat)
                ws.write(i, 1, p_name)
                ws.write(i, 5, c_mat)
                ws.write(i, 6, c_name)
                ws.write(i, 2, 2102)
                ws.write(i, 3, 1)
                if c_mat in self.hibe_mats:
                    ws.write(i, 4, 'N')
                else:
                    ws.write(i, 4, 'L')
                    ws.write(i, 14, 'X')
                    
                ws.write(i, 7, float(value[5]))
                i+=1
                        
        wb.save(file_str)
        logger.info(file_str+'保存完成!')
    
    def generate_app(self):
        if len(self.nstd_mat_list)==0:
            logger.warning('没有非标物料，无法生成非标物料申请表')
            return
        
        nstd_id = simpledialog.askstring('非标申请编号','请输入完整非标申请编号(不区分大小写)，系统将自动关联项目:')
        
        if nstd_id is None:
            return
        
        nstd_id = nstd_id.upper().strip()
        
        basic_info = self.get_rel_nstd_info(nstd_id)
        if not basic_info:
            logger.warning('非标申请：'+nstd_id+'在流程软件中未创建，请先创建后再生成非标物料申请表!')
            return
        
        file_str=filedialog.asksaveasfilename(title="导出文件", initialfile=nstd_id,filetypes=[('excel file','.xlsx')])
        if not file_str:
            return

        if not file_str.endswith(".xlsx"):
            file_str+=".xlsx"
        
        if not self.create_nstd_mat_table(nstd_id, basic_info):
            logger.warning('由于非标物料均已经在其他非标申请中提交，故中止创建非标申请清单文件。')
            return
            
        temp_file = os.path.join(cur_dir(),'temp_eds.xlsx')
        logger.info('正在根据模板文件:'+temp_file+'生成申请表...')
        wb = load_workbook(temp_file) 
        temp_ws = wb.get_sheet_by_name('template')
        m_qty = len(self.nstd_mat_list)
        
        if m_qty%28 ==0:
            s_qty = m_qty/28
        else:
            s_qty = int(m_qty/28)+1
            
        for i in range(1, s_qty+1):
            ws = wb.copy_worksheet(temp_ws)     
            ws.sheet_state ='visible'
            ws.title = 'page'+str(i)
            self.style_worksheet(ws)
            
            ws.cell(row=5,column=1).value = 'Page '+str(i)+'/'+str(s_qty)
            logger.info('正在向第'+str(i)+'页填入物料数据...')
            self.fill_nstd_app_table(ws, i, nstd_id, basic_info,m_qty)
        
        wb.remove_sheet(temp_ws)
        
        self.nstd_app_id = nstd_id
        if writer.excel.save_workbook(workbook=wb, filename=file_str):
            logger.info('生成非标物料申请文件:'+file_str+' 成功!')
        else:
            logger.info('文件保存失败!')
    
    def create_nstd_mat_table(self, nstd_id, res):
        logger.info('正在保存非标物料到维护列表中...')
        no_need_mats=[]
        try:
            nstd_app_head.get(nstd_app_head.nstd_app == nstd_id)
            logger.warning('非标申请:'+nstd_id+'已经存在，故未重新创建!')        
            #q= nstd_app_head.update(project=res['project_id'], contract=res['contract'], index_mat=res['index_mat_id'], app_person=res['app_person']).where(nstd_app_head.nstd_app == nstd_id)
            #q.execute()
        except nstd_app_head.DoesNotExist:
            nstd_app_head.create(nstd_app=nstd_id, project=res['project_id'], contract=res['contract'], index_mat=res['index_mat_id'], app_person=res['app_person'])
        
        wbs_list = res['units']

        for wbs in wbs_list:  
            if  len(wbs.strip())==0 and len(wbs_list)>1:
                continue
            nstd_app_link.get_or_create(nstd_app=nstd_id, wbs_no=wbs, mbom_fin=False) 

        for mat in self.nstd_mat_list:
            line = self.mat_items[mat]
            try:
                r=nstd_app_head.select().join(nstd_mat_table).where(nstd_mat_table.mat_no==mat).naive().get()
                nstd_app = none2str(r.nstd_app)
                logger.error('非标物料:'+mat+'已经在非标申请:'+nstd_app+'中提交，请勿重复提交！')
                if nstd_id != nstd_app and mat not in no_need_mats:
                    no_need_mats.append(mat)
            except nstd_app_head.DoesNotExist:
                rp_sj=''
                box_sj=''
                rp_zs=''
                box_zs=''
                if mat in self.mat_tops.keys():
                    rp_box = self.mat_tops[mat]['rp_box']
                    if rp_box is not None:
                        rp_sj = rp_box['2101'][0]
                        box_sj = rp_box['2101'][1]
                        rp_zs = rp_box['2001'][0]
                        box_zs = rp_box['2001'][1]
                        
                nstd_mat_table.create(mat_no=mat, mat_name_cn=line[mat_heads[2]],
                                      mat_name_en=line[mat_heads[3]], drawing_no=line[mat_heads[4]],
                                      mat_unit=line[mat_heads[6]],comments=line[mat_heads[9]],
                                      rp=rp_sj,box_code_sj=box_sj,rp_zs=rp_zs,box_code_zs=box_zs,
                                      nstd_app = nstd_id, mat_app_person=res['app_person'])

            try:
                nstd_mat_fin.get(nstd_mat_fin.mat_no==mat)
            except nstd_mat_fin.DoesNotExist:
                nstd_mat_fin.create(mat_no=mat,justify=-1, mbom_fin=False,\
                                    pu_price_fin=False, co_run_fin=False, modify_by=login_info['uid'], modify_on=datetime.datetime.now())
        
        for mat in no_need_mats:
            self.nstd_mat_list.remove(mat)
            
        if len(self.nstd_mat_list)==0:
            logger.error(' 所有非标物料已经在另外的非标申请中提交，请勿重复提交!') 
            return False
        else:      
            logger.info('非标物料维护列表保存进程完成.')
            return True
    
    def fill_nstd_app_table(self, ws, page, nstd, res, count):
        ws.cell(row=6, column=2).value = nstd
        ws.cell(row=7, column=4).value = res['project_name']
        ws.cell(row=7, column=20).value= res['contract']
        wbses = res['units']
        ws.cell(row=7, column=12).value = self.combine_wbs(wbses)
        
        if count-count%(page*28)>0:
            ran = 28
        else:
            ran = count%28
        
        for i in range(1, ran+1):
            mat = self.nstd_mat_list[((page-1)*28+i-1)]
            line = self.mat_items[mat]
            ws.cell(row=i+10, column=3).value = line[mat_heads[2]]
            ws.cell(row=i+10, column=4).value = line[mat_heads[3]]
            ws.cell(row=i+10, column=5).value = mat
            drawing_id = line[mat_heads[4]]
            ws.cell(row=i+10, column=7).value = drawing_id
            ws.cell(row=i+10, column=10).value = line[mat_heads[9]]
            ws.cell(row=i+10, column=20).value = line[mat_heads[6]]
            
            if drawing_id=='NO' or len(drawing_id)==0:
                ws.cell(row=i+10, column=21).value ='否'
            else:
                ws.cell(row=i+10, column=21).value = '是'
             
            if mat in self.mat_tops.keys():
                rp_box = self.mat_tops[mat]['rp_box']
                if rp_box is not None:
                    ws.cell(row=i+10, column=15).value = rp_box[login_info['plant']][1]
                    ws.cell(row=i+10, column=17).value = rp_box[login_info['plant']][0]
                                  
    def style_worksheet(self, ws):        
        thin = Side(border_style="thin", color="000000")
        dash = Side(border_style="dashed", color="000000")
                      
        other_border = Border(top=dash, left=dash, right=dash)
        self.set_border(ws, 'T5:V5', other_border)

        main_border = Border(top=thin, left=thin, right=thin, bottom=thin)
        self.set_border(ws, 'A6:V40', main_border)
           
        logo = Image(img=os.path.join(cur_dir(),'logo.png'))
        logo.drawing.top = 0
        logo.drawing.left = 30
        logo.drawing.width=110
        logo.drawing.height=71
        head = Image(img=os.path.join(cur_dir(),'head.png'))
        head.drawing.width = 221
        head.drawing.height = 51
                
        ws.add_image(head,'A2')
        ws.add_image(logo,'T1')
        
        ws.print_area ='A1:V40'
                      
        ws.print_options.horizontalCentered = True
        ws.print_options.verticalCentered = True
        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_margins.left=0.24
        ws.page_margins.right = 0.24
        ws.page_margins.top = 0.19
        ws.page_margins.bottom=0.63
        ws.page_margins.header = 0
        ws.page_margins.footer= 0
        
        ws.page_setup.scale = 80 
        ws.sheet_properties.pageSetUpPr.fitToPage = True 
             
        ws.oddFooter.left.text='''Songjiang Plant,ThyssenKrupp Elevator ( Shanghai ) Co., Ltd.
No.2, Xunye Road, Sheshan Subarea, Songjiang Industrial Area, Shanghai
Tel.: +86 (21) 37869898   Fax: +86 (21) 57793363
TKEC.SJ-F-03-03'''
        ws.oddFooter.left.font='TKTypeMedium, Regular' 
        ws.oddFooter.left.size =7 
               
        ws.oddFooter.right.text='项目非标物料汇总表V2.01'
        ws.oddFooter.right.font='宋体, Regular' 
        #ws.oddFooter.right.size =8          
            
    def set_border(self, ws, cell_range, border): 
        top = Border(top=border.top)
        left = Border(left=border.left)
        right = Border(right=border.right)
        bottom = Border(bottom=border.bottom)               
        rows = ws[cell_range]

        for cell in rows[-1]:
            cell.border = cell.border + bottom 
                   
        for row in rows:
            r = row[-1]
            r.border = r.border+right
            for cell in row:
                cell.border = cell.border+top+left
                                    
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
                                
    def get_rel_nstd_info(self, nstd_id):
        try:
            nstd_result = NonstdAppItem.select(NonstdAppItem.link_list, NonstdAppItemInstance.index_mat, NonstdAppItemInstance.res_engineer, NonstdAppItemInstance.create_emp).join(NonstdAppItemInstance, on=(NonstdAppItem.index == NonstdAppItemInstance.index))\
                .where((NonstdAppItemInstance.nstd_mat_app==nstd_id)&(NonstdAppItem.status>=0)&(NonstdAppItemInstance.status>=0)).naive().get()
        except NonstdAppItem.DoesNotExist:
            return None
        
        res ={}
        
        wbs_res = nstd_result.link_list
        index_mat = nstd_result.index_mat

        try:
            emp_res= SEmployee.get(SEmployee.employee==login_info['uid'])
            app_per = emp_res.name
        except SEmployee.DoesNotExist:
            app_per=''         

        i_pos = index_mat.find('-')
        nstd_app_id = index_mat[0:i_pos]
        try:
            nstd_app_result=NonstdAppHeader.get((NonstdAppHeader.nonstd==nstd_app_id)&(NonstdAppHeader.status>=0))
        except NonstdAppHeader.DoesNotExist:
            return None

        project_id = nstd_app_result.project
        contract_id = nstd_app_result.contract
        
        try:
            p_r = ProjectInfo.get(ProjectInfo.project==project_id)
        except ProjectInfo.DoesNotExist:
            return None
        
        project_name = p_r.project_name
        
        if isinstance(wbs_res, str): 
            wbs_list = wbs_res.split(';')
        else:
            wbs_list=['']
            
        wbses =[]
        
        for wbs in wbs_list:
            if  len(wbs.strip())==0 and len(wbs_list)>1:
                continue
            
            w = wbs.strip()
            w = w[0:14]
            wbses.append(w)
            
        res['units']=wbses
        res['contract']=contract_id
        res['project_id']=project_id
        res['project_name']=project_name
        res['app_person'] = app_per
        res['index_mat_id'] = index_mat
        
        return res
                     
    def run_check_in_sap(self):
        if self.sap_thread is not None:
            if self.sap_thread.is_alive():
                messagebox.showinfo('提示','正在后台检查SAP非标物料，请等待完成后再点击!')
                return            
        
        self.sap_thread = refresh_thread(self)
        self.sap_thread.setDaemon(True)
        self.sap_thread.start()
        
    def refresh(self):
        self.nstd_mat_list=[]
        self.nstd_app_id=''
        self.hibe_mats=[]
                
        logger.info("正在登陆SAP...")
        config = ConfigParser()
        config.read('sapnwrfc.cfg')
        para_conn  = config._sections['connection']
        para_conn['user'] = base64.b64decode(para_conn['user']).decode()
        para_conn['passwd'] = base64.b64decode(para_conn['passwd']).decode()
        mats = self.mat_items.keys()
        
        try:
            conn = pyrfc.Connection(**para_conn)
            
            imp = []
            for mat in mats:
                line = dict(MATNR=mat, WERKS='2101')
                imp.append(line)
            
            logger.info("正在调用RFC函数...")
            result = conn.call('ZAP_PS_MATERIAL_INFO', IT_CE_MARA=imp, CE_SPRAS='1')
            
            std_mats=[]
            for re in result['OT_CE_MARA']:
                std_mats.append(re['MATNR'])
                
                if re['BKLAS']=='3030' and re['MATNR'] not in self.hibe_mats:
                    self.hibe_mats.append(re['MATNR'])
                
            for mat in mats:
                if mat not in std_mats:
                    logger.info("标记非标物料:"+mat)
                    self.nstd_mat_list.append(mat)
                    self.mark_nstd_mat(mat, True)
                else:
                    self.mark_nstd_mat(mat, False)
                    
            logger.info("非标物料确认完成，共计"+str(len(self.nstd_mat_list))+"个非标物料。")
            
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
                   
        return len(self.nstd_mat_list)
        
    def mark_nstd_mat(self, mat, non=True):
        re=mat_info.get(mat_info.mat_no == mat)
        
        if re.is_nonstd == non:
            return 0
        else:
            q = mat_info.update(is_nonstd=non).where(mat_info.mat_no==mat)
            r = q.execute()
            if r>0:
                self.change_log('mat_info', 'is_nonstd',mat , (not non), non)
                
            return r
        
    def multi_find(self):
        d = ask_list('物料拷贝器',2)
        if not d:
            logger.warning('物料清单不能为空，请务必填写物料号')
            return
        
        self.mat_tops = {}        
        self.mat_items={}
        self.mat_list = {}
        self.bom_items = [] 
        self.nstd_mat_list = []
        
        for row in self.mat_tree.get_children():
            self.mat_tree.delete(row) 
        
        logger.info('开始搜索匹配的物料号...')

        res=mat_info.select(mat_info,bom_header.struct_code, bom_header.bom_id, bom_header.revision,bom_header.is_active).join(bom_header, on=(bom_header.mat_no==mat_info.mat_no)).where((mat_info.mat_no.in_(d)) & (bom_header.is_active==True))\
                .order_by(mat_info.mat_no.asc()).naive()  

        if not res:
            logger.warning("没有与搜索条件匹配的物料BOM.")
            return 
        
        self.get_res(res)
        
    def search(self,event=None):
        if len(self.find_var.get())==0:
            logger.warning("物料号不能为空，请务必填写物料号")
            return
        
        self.mat_tops = {}        
        self.mat_items={}
        self.mat_list = {}
        self.bom_items = [] 
        self.nstd_mat_list = []
        
        for row in self.mat_tree.get_children():
            self.mat_tree.delete(row) 
        
        logger.info('开始搜索匹配的物料号...')
        if len(self.version_var.get())==0:
            res=mat_info.select(mat_info,bom_header.struct_code, bom_header.bom_id, bom_header.revision,bom_header.is_active).join(bom_header, on=(bom_header.mat_no==mat_info.mat_no)).where((mat_info.mat_no.contains(self.find_var.get())) & (bom_header.is_active==True))\
                .order_by(mat_info.mat_no.asc()).naive()  
        else:                 
            res=mat_info.select(mat_info, bom_header.bom_id,bom_header.struct_code, bom_header.revision,bom_header.is_active).join(bom_header, on=(bom_header.mat_no==mat_info.mat_no)).where((mat_info.mat_no.contains(self.find_var.get())) & (bom_header.revision==self.version_var.get())& (bom_header.is_active==True))\
                .order_by(mat_info.mat_no.asc()).naive()
            
        if not res:
            logger.warning("没有与搜索条件匹配的物料BOM.")
            return 
        
        self.get_res(res)
    
    def get_res(self, res):    
        for l in res:
            line = {}
            re = {}

            mat = none2str(l.mat_no)
            rev = none2str(l.revision)
                        
            line[mat_heads[0]]= rev
            line[mat_heads[1]]= mat
            line[mat_heads[2]]= none2str(l.mat_name_cn)
            line[mat_heads[3]]= none2str(l.mat_name_en)
            line[mat_heads[4]]= none2str(l.drawing_no)
            line[mat_heads[5]]= 0
            line[mat_heads[6]]= none2str(l.mat_unit)
            line[mat_heads[7]]= none2str(l.mat_material)
            line[mat_heads[8]]= none2str(l.part_weight)
            line[mat_heads[9]]= '' 
            
            #revision = none2str(l.revision)
            #struct_code = none2str(l.struct_code)
            
            re['revision']=none2str(l.revision)
            re['struct_code']=none2str(l.struct_code)
                 
            #if len(struct_code)>0 and mat not in self.mat_tops:
            #re['revision']=revision
            #re['struct_code']=struct_code
            rp_box={}
            if len(none2str(l.rp))!=0 or len(none2str(l.box_code_sj))!=0 or \
                len(none2str(l.rp_zs))!=0 or len(none2str(l.box_code_zs))!=0:
                lt=[]
                lt.append(none2str(l.rp))
                lt.append(none2str(l.box_code_sj))
                rp_box['2101']=lt
                lt=[]                    
                lt.append(none2str(l.rp_zs))
                lt.append(none2str(l.box_code_zs))
                rp_box['2001']=lt
                re['rp_box'] = rp_box
                self.mat_tops[mat]=re
            #else:
            #   rp_box=None
            #re['rp_box'] = rp_box
            #self.mat_tops[mat]=re                
            is_nstd = l.is_nonstd
            if is_nstd and mat not in self.nstd_mat_list:
                self.nstd_mat_list.append(mat)
                    
            if mat not in self.mat_items.keys():
                self.mat_items[mat]=line
            
            item = self.mat_tree.insert('', END, values=dict2list(line))
            
            self.mat_list[item]=line
            if self.get_sub_bom(item, mat, rev):
                self.bom_items.append(item) 
       
        logger.info('正在与SAP匹配确认非标物料，请勿进行其他操作！')
        self.run_check_in_sap()  
    
    def para_search(self, event=None):  
        if len(self.para_var.get())==0:
            logger.warning("参数不能为空，请务必填写 参数")
            return
        
        self.mat_tops = {}        
        self.mat_items={}
        self.mat_list = {}
        self.bom_items = [] 
        self.nstd_mat_list = []
        
        for row in self.mat_tree.get_children():
            self.mat_tree.delete(row) 
        
        logger.info('开始搜索匹配的物料号...')
                
        res=mat_info.select(mat_info, bom_header.bom_id,bom_header.struct_code, bom_header.revision,bom_header.is_active).join(bom_header, on=(bom_header.mat_no==mat_info.mat_no)).where((mat_info.mat_name_cn.contains(self.para_var.get()) | mat_info.comments.contains(self.para_var.get()))& (bom_header.is_active==True))\
                .order_by(mat_info.mat_no.asc()).naive()
            
        if not res:
            logger.warning("没有与搜索条件匹配的物料号.")
            return 
        
        for l in res:
            line = {}
            re = {}

            mat = none2str(l.mat_no)
            rev = none2str(l.revision)
                        
            line[mat_heads[0]]= rev
            line[mat_heads[1]]= mat
            line[mat_heads[2]]= none2str(l.mat_name_cn)
            line[mat_heads[3]]= none2str(l.mat_name_en)
            line[mat_heads[4]]= none2str(l.drawing_no)
            line[mat_heads[5]]= 0
            line[mat_heads[6]]= none2str(l.mat_unit)
            line[mat_heads[7]]= none2str(l.mat_material)
            line[mat_heads[8]]= none2str(l.part_weight)
            line[mat_heads[9]]= '' 
            
            re['revision']=none2str(l.revision)
            re['struct_code']=none2str(l.struct_code)
                 
            rp_box={}
            if len(none2str(l.rp))!=0 or len(none2str(l.box_code_sj))!=0 or \
                len(none2str(l.rp_zs))!=0 or len(none2str(l.box_code_zs))!=0:
                lt=[]
                lt.append(none2str(l.rp))
                lt.append(none2str(l.box_code_sj))
                rp_box['2101']=lt
                lt=[]                    
                lt.append(none2str(l.rp_zs))
                lt.append(none2str(l.box_code_zs))
                rp_box['2001']=lt
                re['rp_box'] = rp_box
                self.mat_tops[mat]=re
                    
            if mat not in self.mat_items.keys():
                self.mat_items[mat]=line
            
            item = self.mat_tree.insert('', END, values=dict2list(line))
             
            self.mat_list[item]=line
            if self.get_sub_bom(item, mat, rev, False):
                self.bom_items.append(item)     
                     
    def get_sub_bom(self,item, mat, rev='', nstd_check=True):
        r = bom_header.select(bom_header, bom_item, mat_info).join(bom_item, on=(bom_header.bom_id==bom_item.bom_id)).switch(bom_item).join(mat_info, on=(bom_item.component==mat_info.mat_no))\
              .where((bom_header.mat_no==mat)&(bom_header.revision==rev)&(bom_header.is_active==True)).order_by(bom_item.index.asc()).naive()
        
        if not r:
            return False
        
        logger.info('开始搜索物料:'+mat+'的下层BOM')
        for l in r:
            line = {}
            re={}
            
            line[mat_heads[0]]= none2str(l.st_no)
            mat = none2str(l.component)
            line[mat_heads[1]]= mat
            line[mat_heads[2]]= none2str(l.mat_name_cn)
            line[mat_heads[3]]= none2str(l.mat_name_en)
            line[mat_heads[4]]= none2str(l.drawing_no)
            line[mat_heads[5]]= l.qty
            line[mat_heads[6]]= none2str(l.mat_unit)
            line[mat_heads[7]]= none2str(l.mat_material)
            line[mat_heads[8]]= none2str(l.part_weight)
            line[mat_heads[9]]= none2str(l.bom_remark)
            
            if nstd_check==True:
                is_nstd = l.is_nonstd
                if is_nstd and mat not in self.nstd_mat_list:
                    self.nstd_mat_list.append(mat)
                
            tree_item = self.mat_tree.insert(item, END, values=dict2list(line))
            self.mat_list[tree_item]=line
                   
            if mat not in self.mat_items.keys():
                self.mat_items[mat]=line
             
            if self.get_sub_bom(tree_item, mat, '', nstd_check):
                self.bom_items.append(tree_item)
        
        logger.info('构建物料:'+mat+'下层BOM完成!')        
        return True
                         
    def check_sub_bom(self, mat, ver=''):
        try:
            bom_header.get((bom_header.mat_no==mat)&(bom_header.revision==ver))
            return True
        except bom_header.DoesNotExist:
            return False        
                                         
    def excel_import(self):
        file_list = filedialog.askopenfilenames(title="导入文件", filetypes=[('excel file','.xlsx'),('excel file','.xlsm')])
        if not file_list:
            return

        self.mat_list = {}
        self.mat_pos = 0
        self.mat_tops = {}
        
        self.mat_items={}
        
        for row in self.mat_tree.get_children():
            self.mat_tree.delete(row)
             
        #for node in self.bom_tree.children(0):
            #self.bom_tree.remove_node(node.identifier)

        for file in file_list:
            logger.info("正在读取文件:"+file+",转换保存物料信息,同时构建数据Model")
            c=self.read_excel_files(file)
            logger.info("文件:"+file+"读取完成, 共计处理 "+str(c)+" 个物料。")
            
        #df = pd.DataFrame(self.mat_list,index=mat_heads, columns=[ i for i in range(1, self.mat_pos+1)])
        #model = TableModel(dataframe=df.T)
        #self.mat_table.updateModel(model)
        #self.mat_table.redraw()
        
        logger.info("正在生成BOM层次结构...")
        c = self.build_tree_struct()
        logger.info("Bom结构生成完成，共为"+str(c)+"个发运层物料生成BOM.")
        
        logger.info("正在保存BOM...")
        c = self.save_mats_bom()
        logger.info("共保存"+str(c)+"个物料BOM")
        
        logger.info("正在核查非标物料...")
        self.run_check_in_sap()
                      
    def save_mat_info(self,method=False,**para):
        b_level=False
        
        if para['mat_no'] in self.mat_tops.keys():
            rp_box=  self.mat_tops[para['mat_no']]['rp_box']
            if rp_box is not None:
                b_level=True
            
        try:
            mat_info.get(mat_info.mat_no == para['mat_no'])
            if method:
                if b_level:
                    q = mat_info.update(mat_name_en=para['mat_name_en'], mat_name_cn=para['mat_name_cn'], drawing_no=para['drawing_no'],mat_material=para['mat_material'],mat_unit=para['mat_unit'],rp=rp_box['2101'][0], box_code_sj=rp_box['2101'][1],\
                               rp_zs=rp_box['2001'][0],box_code_zs=rp_box['2001'][1],mat_material_en=para['mat_material_en'],part_weight=para['part_weight'],comments=para['comments'],modify_by=login_info['uid'],modify_on=datetime.datetime.now()).where(mat_info.mat_no==para['mat_no']) 
                else:
                    q = mat_info.update(mat_name_en=para['mat_name_en'], mat_name_cn=para['mat_name_cn'], drawing_no=para['drawing_no'],mat_material=para['mat_material'],mat_unit=para['mat_unit'],\
                                mat_material_en=para['mat_material_en'],part_weight=para['part_weight'],comments=para['comments'],modify_by=login_info['uid'],modify_on=datetime.datetime.now()).where(mat_info.mat_no==para['mat_no'])
                return q.execute()                          
        except mat_info.DoesNotExist:
            if b_level:
                q = mat_info.insert(mat_no=para['mat_no'], mat_name_en=para['mat_name_en'], mat_name_cn=para['mat_name_cn'], drawing_no=para['drawing_no'],mat_material=para['mat_material'],mat_unit=para['mat_unit'],\
                                mat_material_en=para['mat_material_en'],part_weight=para['part_weight'],rp=rp_box['2101'][0],box_code_sj=rp_box['2101'][1],rp_zs=rp_box['2001'][0],box_code_zs=rp_box['2001'][1],comments=para['comments'],modify_by=login_info['uid'],modify_on=datetime.datetime.now())
            else:
                q = mat_info.insert(mat_no=para['mat_no'], mat_name_en=para['mat_name_en'], mat_name_cn=para['mat_name_cn'], drawing_no=para['drawing_no'],mat_material=para['mat_material'],mat_unit=para['mat_unit'],\
                                mat_material_en=para['mat_material_en'],part_weight=para['part_weight'],comments=para['comments'],modify_by=login_info['uid'],modify_on=datetime.datetime.now())
            return q.execute()
        
        return 0
    
    def check_branch(self, item):
        mat = self.mat_tree.item(item, "values")[1]
        for li in self.bom_items:
            if mat == self.mat_tree.item(li, "values")[1]:
                return False
            
        self.bom_items.append(item)
        
        return True
        
    def save_bom_list(self, item):
        it_list = self.mat_tree.item(item, "values")
        mat = it_list[1]
        drawing = it_list[4]
        
        if mat in self.mat_tops.keys():
            revision = self.mat_tops[mat]['revision']
            st_code = self.mat_tops[mat]['struct_code']
        else:
            revision=''
            st_code=''
        
        try:    
            bom_header.get((bom_header.mat_no==mat) & (bom_header.revision==revision) & (bom_header.is_active==True))
            logger.warning(mat+"BOM已经存在，无需重新创建!")
            return 0
        except bom_header.DoesNotExist:
            b_id = self.bom_id_generator()
            q=bom_header.insert(bom_id=b_id, mat_no=mat, revision=revision, drawing_no=drawing, struct_code=st_code,is_active=True,plant=login_info['plant'],\
                                modify_by=login_info['uid'],modify_on=datetime.datetime.now(), create_by=login_info['uid'],create_on=datetime.datetime.now())
            q.execute()
            
        children = self.mat_tree.get_children(item)
        
        data = []
        for child in children:
            d_line = {}
            d_line['bom_id']=b_id
            d_line['index'] = int(self.mat_tree.item(child, "values")[0])
            d_line['st_no'] = self.mat_tree.item(child, "values")[0]
            d_line['component'] = self.mat_tree.item(child,"values")[1]
            d_line['qty']= Decimal(self.mat_tree.item(child,"values")[5])
            d_line['bom_remark']=self.mat_tree.item(child,"values")[9]
            d_line['parent_mat'] = mat
            d_line['modify_by']=login_info['uid']
            d_line['modify_on']= datetime.datetime.now()
            d_line['create_by']=login_info['uid']
            d_line['create_on']=datetime.datetime.now()
            
            data.append(d_line)
            
        q=bom_item.insert_many(data)
        return q.execute()
                        
    def get_rp_boxid(self, struct, plant='2101'):
        rp_box = {}
               
        res=struct_gc_rel.select().where(struct_gc_rel.st_code==struct)
           
        for r in res:
            lt=[]
            lt.append(r.rp)
            lt.append(r.box_code)
            rp_box[r.plant]= lt
      
        return rp_box
    
    def save_mats_bom(self):
        if len(self.bom_items)==0:
            return 0
        
        i=0
        for item in self.bom_items:
            if self.save_bom_list(item)>0:
                i+=1
        
        return i    
    
    def build_tree_struct(self):
        self.bom_items=[]
        if len(self.mat_list)==0:
            return 0
        
        cur_level = 0
        pre_level = 0
        parent_node = self.mat_tree.insert('', END, values = dict2list(self.mat_list[1]))
        counter =0
        cur_node = parent_node
        self.check_branch(parent_node)
        
        self.mat_tree.item(parent_node, open=True)
        
        for i in range(1,self.mat_pos+1):
            cur_level = tree_level(self.mat_list[i][mat_heads[0]])
            if cur_level==0:
                counter+=1
                
            if (pre_level == cur_level) and pre_level !=0:
                cur_node = self.mat_tree.insert(parent_node, END,  values=dict2list(self.mat_list[i]))
                
            if pre_level<cur_level:
                parent_node = cur_node
                self.check_branch(parent_node)
                cur_node=self.mat_tree.insert(parent_node, END, values=dict2list(self.mat_list[i]))
                
            if pre_level>cur_level:
                while pre_level >= cur_level:
                    parent_node = self.mat_tree.parent(parent_node)
                    if pre_level!=0:
                        pre_level = tree_level(self.mat_tree.item(parent_node, 'values')[0])
                    else:
                        pre_level=-1
                    
                cur_node=self.mat_tree.insert(parent_node, END, values=dict2list(self.mat_list[i]))
                
                if cur_level==0:
                    self.mat_tree.item(cur_node, open=True)
                
            pre_level = cur_level
                
        return counter

    '''        
    def build_tree_struct(self):
        if len(self.mat_list)==0:
            return
        
        cur_level=0
        pre_level=0
        parent_node=0
        counter=0
        for i in range(1, self.mat_pos+1):
            cur_level = tree_level(self.mat_list[i][mat_heads[0]])
            if cur_level==0:
                counter+=1
                
            if pre_level == cur_level:
                self.bom_tree.create_node(i,i,parent_node)
                
            if pre_level < cur_level:
                parent_node = i-1
                self.bom_tree.create_node(i,i,parent_node)
                
            if pre_level > cur_level:
                while pre_level > cur_level:
                    parent_node = self.bom_tree.parent(parent_node).identifier
                    pre_level = tree_level(self.mat_list[parent_node][mat_heads[0]])
                    
                self.bom_tree.create_node(i,i,parent_node)
                                 
            pre_level = cur_level
        
        return counter
    ''' 
                     
    def read_excel_files(self, file):
        '''
                返回值：
                -2: 读取EXCEL失败
                -1 : 头物料位置为空
                0： 头物料的版本已经存在
                1： 
       
        '''
        wb = load_workbook(file, read_only=True,data_only=True)
        sheetnames=wb.get_sheet_names()
        
        if len(sheetnames)==0:
            return -2
        
        counter=0
    
        for i in range(0,len(sheetnames)): 
            if not str(sheetnames[i]).isdigit():
                continue
            
            for j in range(1,19):                         
                mat_line = {}
                mat_top_line={}
                mat=''
                ws = wb.get_sheet_by_name(sheetnames[i]) 
            
                mat_line[mat_heads[0]]=cell2str(ws.cell(row=2*j+1,column=2).value)
                mat = cell2str(ws.cell(row=2*j+1,column=5).value)
                
                if len(mat)==0:
                    break
                
                mat_line[mat_heads[1]]= mat        
                mat_line[mat_heads[2]]=cell2str(ws.cell(row=2*j+1,column=7).value)
                mat_line[mat_heads[3]]=cell2str(ws.cell(row=2*j+2,column=7).value)
                mat_line[mat_heads[4]] = cell2str(ws.cell(row=2*j+1, column=6).value)
                
                qty = cell2str(ws.cell(row=2*j+1,column=3).value)
                if len(qty)==0: 
                    continue
                
                self.mat_pos+=1
                counter+=1
                
                mat_line[mat_heads[5]] = Decimal(qty)
                mat_line[mat_heads[6]]=cell2str(ws.cell(row=2*j+1,column=4).value)
                mat_line[mat_heads[7]] = cell2str(ws.cell(row=2*j+1,column=9).value)
                material_en = cell2str(ws.cell(row=2*j+2, column=9).value)
                
                weight = cell2str(ws.cell(row=2*j+1, column=10).value)
                if len(weight)==0:
                    mat_line[mat_heads[8]]=0
                else:
                    mat_line[mat_heads[8]]=Decimal(weight)
                    
                mat_line[mat_heads[9]]=cell2str(ws.cell(row=2*j+1, column=11).value)
                
                len_of_st = len(mat_line[mat_heads[0]])
                str_code = cell2str(ws.cell(row=39,column=12).value)
                if len_of_st <=1:
                    if len_of_st==0:
                        mat_top_line['revision'] = cell2str(ws.cell(row=43, column=8).value)
                        mat_top_line['struct_code']=str_code
                    else:
                        mat_top_line['revision'] = ''
                        mat_top_line['struct_code']=''
                                            
                    rp_box = self.get_rp_boxid(str_code)
                    mat_top_line['rp_box'] = rp_box
                    
                    self.mat_tops[mat_line[mat_heads[1]]]=mat_top_line
                
                #保存物料基本信息
                if self.save_mat_info(mat_no=mat_line[mat_heads[1]], mat_name_en=mat_line[mat_heads[3]], mat_name_cn=mat_line[mat_heads[2]], drawing_no=mat_line[mat_heads[4]],mat_material=mat_line[mat_heads[7]],mat_unit=mat_line[mat_heads[6]],\
                                mat_material_en=material_en,part_weight=mat_line[mat_heads[8]],comments=mat_line[mat_heads[9]])==0:
                    logger.info(mat_line[mat_heads[1]]+'数据库中已经存在,故没有保存')
                else:
                    logger.info(mat_line[mat_heads[1]]+'保存成功。')
                
                self.mat_list[self.mat_pos] = mat_line
                
                if mat not in self.mat_items.keys():
                    self.mat_items[mat] = mat_line
                               
        return counter
                                                    
    def bom_id_generator(self):
        try:
            bom_res = id_generator.get(id_generator.id == 1)
        except id_generator.DoesNotExist:
            return None
        
        pre_char = none2str(bom_res.pre_character)
        fol_char = none2str(bom_res.fol_character)
        c_len = bom_res.id_length
        cur_id = bom_res.current
        step = bom_res.step        
        new_id=str(cur_id+step)
        #前缀+前侧补零后长度为c_len+后缀, 组成新的BOM id               
        id_char = pre_char+new_id.zfill(c_len)+fol_char
        
        q=id_generator.update(current=cur_id+step).where(id_generator.id==1)
        q.execute()
        
        return id_char
    
    def change_log(self,table,section,key, old,new):
        q = s_change_log.insert(table_name=table,change_section=section,key_word=str(key),old_value=str(old),new_value=str(new),log_on=datetime.datetime.now(), log_by=login_info['uid'] )
        q.execute()   
        
        
#分箱程序
class packing_pane(Frame):
    def __init__(self,master=None):
        Frame.__init__(self, master)
        self.grid()
        
        self.createWidgets()
        
    def createWidgets(self):
        pass               