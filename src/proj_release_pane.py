#coding=utf-8
'''
Created on 2017年1月24日

@author: 10256603
'''
from global_list import *
global login_info

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
