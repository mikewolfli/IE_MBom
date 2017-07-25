# coding=utf-8
'''
Created on 2017年1月24日

@author: 10256603
'''
from global_list import *
global login_info


display_col = ['col0', 'col1', 'col2', 'col3', 'col4', 'col5', 'col6', 'col7', 'col8',
               'col9', 'col10', 'col11', 'col21', 'col22', 'col23', 'col24', 'col25', 'col26','col27']
cols = ['col0', 'col1', 'col2', 'col3', 'col4', 'col5', 'col6', 'col7', 'col8', 'col9', 'col10', 'col11', 'col12', 'col13',
        'col14', 'col15', 'col16', 'col17', 'col18', 'col19', 'col20', 'col21', 'col22', 'col23', 'col24', 'col25', 'col26','col27']
tree_head = ['导入日期', '非标编号', '判断', '物料号', '物料名称(中)', '物料名称(英)', '图号', '单位', '备注', 'RP', 'BoxId', '申请人', '自制完成', '负责人',
             '完成日期', '价格维护', '负责人', '完成日期', 'CO-run', '负责人', '完成日期', '是否有图纸', '图纸上传', 'IN SAP', '是否紧急', '项目交货期', '要求配置完成日期','物料要求维护完成日期']


class mat_fin_pane(Frame):
    '''
    所有符合条件的数据存储在mat_res这个字典中
    字典结构：
     mat_res-{mat_no：[第一级value1],...}
     nstd_res-{nstd: [wbs_no_list],...}
     wbs_res-{wbs_no: [unit_info],....}

    '''
    data_thread = None
    b_finished = False
    mat_res = {}
    nstd_res = {}
    # 存储unit wbs相关信息， key-wbs no, value-列表[wbs no, 合同号， 项目名称，
    # 梯号，梯型，载重，速度，非标类型，是否紧急，项目交货期，要求配置完成日期]
    wbs_res = {}
    # 存储部分计算值， key=nstd_app, value-列表[图纸上传，in SAP, 是否紧急， 项目交货期，要求配置完成日期]
    nstd_temp_res = {}

    def __init__(self, master=None):
        Frame.__init__(self, master)

        box_id_head = Label(self, text='箱号', anchor='w')
        box_id_head.grid(row=0, column=0, sticky=W)

        self.boxid_var = StringVar()
        self.boxid_cb = ttk.Combobox(
            self, textvariable=self.boxid_var, values=self.get_list_for(0), state='readonly')
        self.boxid_cb.grid(row=1, column=0, sticky=EW)
        self.boxid_var.set('All')
        self.boxid_cb.bind('<<ComboboxSelected>>', self.filter_list)

        rp_head = Label(self, text='RP', anchor='w')
        rp_head.grid(row=0, column=1, sticky=W)
        self.rp_var = StringVar()
        self.rp_cb = ttk.Combobox(
            self, textvariable=self.rp_var, values=self.get_list_for(1), state='readonly')
        self.rp_var.set('All')
        self.rp_cb.grid(row=1, column=1, sticky=EW)
        self.rp_cb.bind('<<ComboboxSelected>>', self.filter_list)

        mat_head = Label(self, text='物料号搜索', anchor='w')
        mat_head.grid(row=0, column=2, sticky=W)
        self.mat_var = StringVar()
        self.mat_entry = Entry(self, textvariable=self.mat_var)
        self.mat_entry.grid(row=1, column=2, sticky=EW)
        self.mat_entry.bind('<KeyRelease>', self.valid_mat)
        self.mat_entry.bind("<Return>", self.mat_search)

        self.elec_mt_button = Button(self, text='电气完成')
        self.elec_mt_button.grid(row=0, column=3, sticky=EW)
        self.elec_mt_button['command'] = self.elec_fin
        if login_info['perm'][1] != '2' and login_info['perm'][1] != '9':
            self.elec_mt_button.grid_forget()

        self.metal_mt_button = Button(self, text='钣金完成')
        self.metal_mt_button.grid(row=0, column=4, sticky=EW)
        self.metal_mt_button['command'] = self.metal_fin
        if login_info['perm'][1] != '3' and login_info['perm'][1] != '9':
            self.metal_mt_button.grid_forget()

        self.price_button = Button(self, text='价格维护完成')
        self.price_button.grid(row=1, column=3, sticky=EW)
        self.price_button['command'] = self.price_fin
        if login_info['perm'][1] != '4' and login_info['perm'][1] != '9':
            self.price_button.grid_forget()
            

        self.co_run_button = Button(self, text='CO-RUN完成')
        self.co_run_button.grid(row=1, column=4, sticky=EW)
        self.co_run_button['command'] = self.co_run_fin
        if login_info['perm'][1] != '5' and login_info['perm'][1] != '9':
            self.co_run_button.grid_forget()

        self.tm_mt_button = Button(self, text='曳引机完成')
        self.tm_mt_button.grid(row=0, column=5, sticky=EW)
        self.tm_mt_button['command'] = self.traction_fin
        if login_info['perm'][1] != '6' and login_info['perm'][1] != '9':
            self.tm_mt_button.grid_forget()

        self.refresh_button = Button(self, text='手动刷新')
        self.refresh_button.grid(row=1, column=6, sticky=EW)
        self.refresh_button['command'] = self.refresh_by_hand

        self.export_button = Button(self, text='导出EXCEL')
        self.export_button.grid(row=0, column=6,  sticky=EW)
        self.export_button['command'] = self.__export_excel

        self.export_direct = Button(self, text='直接导出已完成清单\n(按CO run完成日期筛选)')
        self.export_direct.grid(row=0, column=7, sticky=EW)
        self.export_direct['command'] = self.direct_export
        if login_info['perm'][1] != '1' and login_info['perm'][1] != '9':
            self.export_direct.grid_forget()

        self.export_finished = Button(self, text='已完成清单\n(按CO run完成日期筛选)')
        self.export_finished.grid(row=1, column=7, sticky=EW)
        self.export_finished['command'] = self.get_finished
        if login_info['perm'][1] != '1' and login_info['perm'][1] != '9':
            self.export_finished.grid_forget()

        if login_info['perm'][1] == '1' or login_info['perm'][1] == '5' or login_info['perm'][1] == '9':
            display_col.insert(-6, 'col12')
            display_col.insert(-6, 'col13')
            display_col.insert(-6, 'col14')
        if login_info['perm'][1] == '1' or login_info['perm'][1] == '5' or login_info['perm'][1] == '9':
            display_col.insert(-6, 'col15')
            display_col.insert(-6, 'col16')
            display_col.insert(-6, 'col17')
        if login_info['perm'][1] == '1' or login_info['perm'][1] == '9':
            display_col.insert(-6, 'col18')
            display_col.insert(-6, 'col19')
            display_col.insert(-6, 'col20')

        self.mat_list = ttk.Treeview(self, columns=cols, displaycolumns=display_col,
                                     selectmode='extended')
        style = ttk.Style()
        style.configure("Treeview", font=('TkDefaultFont', 10))
        style.configure("Treeview.Heading", font=('TkDefaultFont', 9))
        self.mat_list.heading("#0", text='')
        for col in cols:
            i = cols.index(col)
            #self.mat_list.heading(col, text=tree_head[i])
            self.mat_list.heading(col, text=tree_head[
                                  i], command=lambda _col=col: treeview_sort_column(self.mat_list, _col, False))
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
        self.mat_list.column('col27', width=100, anchor='w')
        ysb = ttk.Scrollbar(self, orient='vertical',
                            command=self.mat_list.yview)
        xsb = ttk.Scrollbar(self, orient='horizontal',
                            command=self.mat_list.xview)
        self.mat_list.grid(row=2, column=0, rowspan=2,
                           columnspan=9, sticky='nsew')
        
    
        if login_info['perm'][1] == '4' or  login_info['perm'][1] == '9':
            self.popup=Menu(self, tearoff=0)
            self.popup.add_command(label="显示正在CO-RUN物料", command=self.co_run_display)
            
            self.mat_list.bind("<Button-3>", self.do_popup)
          

        ysb.grid(row=2, column=9, rowspan=2, sticky='ns')
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

        if login_info['status'] and int(login_info['perm'][1]) >= 1:
            self.__loop_refresh()


    def do_popup(self, event):
        # display the popup menu
        try:
            self.popup.tk_popup(event.x_root, event.y_root, 0)
        finally:
            # make sure to release the grab (Tk 8.0a1 only)
            self.popup.grab_release()
            
    def co_run_display(self):
        his_display(self,'非标物料',None, 6)
            
    def combine_wbs(self, li):
        li.sort()
        if len(li) > 1:
            head = li[0]
        elif li is None:
            return ''
        elif len(li) == 0:
            return ''
        else:
            return li[0]

        start = int(li[0][11:])
        j = 1
        end = ''
        for i in range(1, len(li)):
            if int(li[i][11:]) == start + j:
                j += 1
            else:
                if j > 1:
                    head = head + '~' + end
                elif len(end) > 0:
                    head = head + ',' + end

                if j > 1:
                    head = head + ',' + li[i][11:]
                start = int(li[i][11:])

                j = 1
            end = li[i][11:]

        if j > 1:
            head = head + '~' + end
        else:
            head = head + ',' + end

        return head

    def __export_excel(self):
        if self.data_thread.is_alive():
            messagebox.showinfo('提示', '表单刷新线程正在后台刷新列表，请等待完成后再点击!')
            return

        items = self.mat_list.get_children('')
        if not items:
            return

        file_str = filedialog.asksaveasfilename(
            title="导出文件", filetypes=[('excel file', '.xlsx')])
        if not file_str:
            return

        if not file_str.endswith(".xlsx"):
            file_str += ".xlsx"

        wb = Workbook()
        ws = wb.worksheets[0]
        ws.title = '物料清单'
        col_size = len(cols)
        for i in range(col_size):
            ws.cell(row=1, column=i + 1).value = tree_head[i]
        ws.cell(row=1, column=col_size + 1).value = '关联WBS'
        ws.cell(row=1, column=col_size + 2).value = '合同号'
        ws.cell(row=1, column=col_size + 3).value = '项目名称'
        ws.cell(row=1, column=col_size + 4).value = '关联梯台数'
        n = 0

        dic_str = {}
        for item in items:
            if self.mat_list.parent(item) != '':
                continue

            for i in range(col_size):
                ws.cell(row=n + 2, column=i +
                        1).value = self.mat_list.item(item, 'values')[i]

            j = cols.index('col1')
            nstd_id = self.mat_list.item(item, 'values')[j]

            if nstd_id not in dic_str:
                s_str = self.combine_wbs(self.nstd_res[nstd_id])
                dic_str[nstd_id] = s_str

            ws.cell(row=n + 2, column=col_size + 1).value = dic_str[nstd_id]
            if dic_str[nstd_id].startswith('E/'):
                wbs = self.nstd_res[nstd_id][0]
                prj_info = self.wbs_res[wbs]
                ws.cell(row=n + 2, column=col_size + 2).value = prj_info[1]
                ws.cell(row=n + 2, column=col_size + 3).value = prj_info[2]

            ws.cell(row=n + 2, column=col_size +
                    4).value = len(self.nstd_res[nstd_id])
            n += 1

        if excel_xlsx.save_workbook(workbook=wb, filename=file_str):
            messagebox.showinfo("输出", "成功输出!")

    def direct_export(self):
        date_ctrl = date_picker(self)
        self.sel_range = date_ctrl.result
        if not self.sel_range:
            return

        file_str = filedialog.asksaveasfilename(
            title="导出文件", filetypes=[('excel file', '.xlsx')])
        if not file_str:
            return

        if not file_str.endswith(".xlsx"):
            file_str += ".xlsx"

        self.b_finished = True
        self.refresh_mat()
        self.b_finished = False

        wb = Workbook()
        ws = wb.worksheets[0]
        ws.title = '物料清单'
        col_size = len(cols)
        for i in range(col_size):
            ws.cell(row=1, column=i + 1).value = tree_head[i]
        ws.cell(row=1, column=col_size + 1).value = '关联WBS'
        ws.cell(row=1, column=col_size + 2).value = '合同号'
        ws.cell(row=1, column=col_size + 3).value = '项目名称'
        n = 0

        dic_str = {}

        for key in self.mat_res.keys():
            for i in range(col_size):
                ws.cell(row=n + 2, column=i + 1).value = self.mat_res[key][i]

            j = cols.index('col1')
            nstd_id = self.mat_res[key][j]

            if nstd_id not in dic_str:
                s_str = self.combine_wbs(self.nstd_res[nstd_id])
                dic_str[nstd_id] = s_str

            ws.cell(row=n + 2, column=col_size + 1).value = dic_str[nstd_id]
            if dic_str[nstd_id].startswith('E/'):
                wbs = self.nstd_res[nstd_id][0]
                prj_info = self.wbs_res[wbs]
                ws.cell(row=n + 2, column=col_size + 2).value = prj_info[1]
                ws.cell(row=n + 2, column=col_size + 3).value = prj_info[2]
            n += 1

        if excel_xlsx.save_workbook(workbook=wb, filename=file_str):
            messagebox.showinfo("输出", "成功输出!")

    def get_finished(self):
        if self.data_thread.is_alive():
            messagebox.showinfo('提示', '表单刷新线程正在后台刷新列表，请等待完成后再点击!')
            return

        date_ctrl = date_picker(self)
        self.sel_range = date_ctrl.result
        if not self.sel_range:
            return

        self.b_finished = True
        self.__refresh_tree()
        self.b_finished = False

    def metal_fin(self):
        # self.choice=1
        self.__mat_update(1)

    def elec_fin(self):
        # self.choice=1
        self.__mat_update(1)

    def price_fin(self):
        # self.choice=2
        self.__mat_update(2)

    def co_run_fin(self):
        # self.choice=3
        self.__mat_update(3)

    def traction_fin(self):
        # self.choice=1
        self.__mat_update(1)

    def filter_list(self, event):
        rp = self.rp_var.get()
        boxid = self.boxid_var.get()
        self.update_tree_data(col9=rp, col10=boxid)

    def refresh(self):
        self.rp_var.set('All')
        self.boxid_var.set('All')
        self.refresh_mat()
        self.update_tree_data()

    def get_list_for(self, fu):
        if fu == 0:
            fd = fn.Distinct(nstd_mat_table.box_code_sj)
        if fu == 1:
            fd = fn.Distinct(nstd_mat_table.rp)

        res = nstd_mat_table.select(fd)
        all_tub = ['All']

        for r in res:
            if fu == 0:
                all_tub.append(r.box_code_sj)
            if fu == 1:
                all_tub.append(r.rp)

        return all_tub

    def get_name(self, pid):
        if pid == '' or not pid:
            return ''

        try:
            r_name = s_employee.get(s_employee.employee == pid)
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

        self.mat_res = {}
        self.wbs_res = {}
        self.nstd_res = {}
        self.nstd_temp_res = {}

        if login_info['perm'][1] == '2':
            mat_items = nstd_app_head.select(nstd_app_head, nstd_mat_table, nstd_mat_fin).join(nstd_mat_table).switch(nstd_mat_table).join(
                nstd_mat_fin).where((nstd_mat_fin.justify == 3) & (nstd_mat_fin.mbom_fin == False) & (nstd_mat_fin.co_run_fin == False)).naive()
        elif login_info['perm'][1] == '3':
            mat_items = nstd_app_head.select(nstd_app_head, nstd_mat_table, nstd_mat_fin).join(nstd_mat_table).switch(nstd_mat_table).join(
                nstd_mat_fin).where((nstd_mat_fin.justify == 4) & (nstd_mat_fin.mbom_fin == False) & (nstd_mat_fin.co_run_fin == False)).naive()
        elif login_info['perm'][1] == '4':
            mat_items = nstd_app_head.select(nstd_app_head, nstd_mat_table, nstd_mat_fin).join(nstd_mat_table).switch(nstd_mat_table).join(nstd_mat_fin).where(
                ((nstd_mat_fin.justify == 1) | (nstd_mat_fin.justify == 2) | (nstd_mat_fin.justify == 6)) & (nstd_mat_fin.pu_price_fin == False) & (nstd_mat_fin.co_run_fin == False)).naive()
        elif login_info['perm'][1] == '5':
            mat_items = nstd_app_head.select(nstd_app_head, nstd_mat_table, nstd_mat_fin).join(nstd_mat_table).switch(nstd_mat_table).join(nstd_mat_fin)\
                .where(((((nstd_mat_fin.justify == 1) | (nstd_mat_fin.justify == 2) | (nstd_mat_fin.justify == 6)) & (nstd_mat_fin.pu_price_fin == True)) | (((nstd_mat_fin.justify == 3) | (nstd_mat_fin.justify == 4) | (nstd_mat_fin.justify == 5)) & (nstd_mat_fin.mbom_fin == True))) & (nstd_mat_fin.co_run_fin == False)).naive()
        elif login_info['perm'][1] == '6':
            mat_items = nstd_app_head.select(nstd_app_head, nstd_mat_table, nstd_mat_fin).join(nstd_mat_table).switch(nstd_mat_table).join(
                nstd_mat_fin).where((nstd_mat_fin.justify == 5) & (nstd_mat_fin.mbom_fin == False) & (nstd_mat_fin.co_run_fin == False)).naive()
        elif login_info['perm'][1] == '1' or login_info['perm'][1] == '9':
            if not self.b_finished:
                mat_items = nstd_app_head.select(nstd_app_head, nstd_mat_table, nstd_mat_fin).join(nstd_mat_table).switch(
                    nstd_mat_table).join(nstd_mat_fin).where((nstd_mat_fin.justify >= 0) & (nstd_mat_fin.co_run_fin == False)).naive()
            else:
                mat_items = nstd_app_head.select(nstd_app_head, nstd_mat_table, nstd_mat_fin).join(nstd_mat_table).switch(nstd_mat_table).join(nstd_mat_fin).where((nstd_mat_fin.justify >= 0) & (
                    nstd_mat_fin.co_run_fin == True) & (nstd_mat_fin.co_run_fin_on >= self.sel_range['from']) & (nstd_mat_fin.co_run_fin_on <= self.sel_range['to'])).naive()
        else:
            return False

        if not mat_items:
            return False

        self.build_data_model(mat_items)

        return True

    def build_data_model(self, query_res):# fix bug cs/ie sap finish date
        i = 0
        
        for r in query_res:
            item = []
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
            elif len(app_per) == 0:
                app_per = r.mat_app_person

            item.append(app_per)
            temp = (r.mbom_fin and 'Y' or '')
            item.append(temp)
            if temp.upper() == 'Y':
                item.append(get_name(r.mbom_fin_by))
                item.append(none2str(r.mbom_fin_on))
            else:
                item.append('')
                item.append('')
            temp = (r.pu_price_fin and 'Y' or '')
            item.append(temp)
            if temp.upper() == 'Y':
                item.append(get_name(r.pu_price_fin_by))
                item.append(none2str(r.pu_price_fin_on))
            else:
                item.append('')
                item.append('')
                
            temp = (r.co_run_fin and 'Y' or '')
            item.append(temp)
            if temp.upper() == 'Y':
                item.append(get_name(r.co_run_fin_by))
                item.append(none2str(r.co_run_fin_on))
            else:
                item.append('')
                item.append('')

            if not r.drawing_no or r.drawing_no == '':
                has_drawing = ''
            else:
                has_drawing = 'Y'
            item.append(has_drawing)

            wbs_list = self.get_wbs_list(nstd_app_id)

            if nstd_app_id not in self.nstd_temp_res:
                self.nstd_temp_res[nstd_app_id] = self.get_nstd_info(index_mat_id, wbs_list)

            res = self.nstd_temp_res[nstd_app_id]
            for j in range(len(res)):
                item.append(none2str(res[j]))
            
            item.append(datetime2str(r.req_fin_on))
            self.mat_res[i] = item
            i += 1

    def get_nstd_info(self, index, wbs_list):
        if index == '' or not index:
            return ''

        item = []
        try:
            result = LProcAct.select(LProcAct.finish_date).where(
                (LProcAct.instance == index) & (LProcAct.action == 'AT00000015')).get()
            s_date = result.finish_date
        except LProcAct.DoesNotExist:
            s_date = ''

        item.append(s_date)

        try:
            result = LProcAct.select(LProcAct.finish_date).where(
                (LProcAct.instance == index) & (LProcAct.action == 'AT00000020')).get()
            s_date = result.finish_date
        except LProcAct.DoesNotExist:
            s_date = ''

        item.append(s_date)
        req_conf_fin = None
        req_delivery = None
        urgent_flag = ''

        for wbs_line in wbs_list:
            if len(wbs_line) >= 20:
                if wbs_line[23] == 'Y':
                    urgent_flag = 'Y'

                if wbs_line[25] and req_conf_fin is None:
                    req_conf_fin = wbs_line[25]
                elif wbs_line[25] and req_conf_fin:
                    if wbs_line[25] < req_conf_fin:
                        req_conf_fin = wbs_line[25]

                if wbs_line[24] and req_delivery is None:
                    req_delivery = wbs_line[24]
                elif wbs_line[24] and req_delivery:
                    if wbs_line[24] < req_delivery:
                        req_delivery = wbs_line[24]

        item.append(urgent_flag)
        item.append(none2str(req_delivery))
        item.append(none2str(req_conf_fin))

        return item

    def get_wbs_list(self, nstd):
        """

        :type nstd: object
        """
        result = nstd_app_head.select(nstd_app_head, nstd_app_link.wbs_no).join(
            nstd_app_link).where(nstd_app_head.nstd_app == nstd).order_by(nstd_app_link.wbs_no).naive()
        wbs_list = []
        wbs_res_list = []
        for r in result:
            wbs_no = r.wbs_no
            if not isinstance(wbs_no, str) or (isinstance(wbs_no, str) and wbs_no == ''):
                wbs_no = r.project
            if wbs_no not in self.wbs_res:
                temp = self.get_wbs_info(wbs_no)
                if temp:
                    self.wbs_res[wbs_no] = temp
            wbs_list.append(wbs_no)

            wbs_res_list.append(self.wbs_res[wbs_no])

        if nstd not in self.nstd_res:
            self.nstd_res[nstd] = wbs_list

        return wbs_res_list

    def get_wbs_info(self, wbs):
        wbs_info = []
        if not wbs.startswith('E'):
            try:
                result = ProjectInfo.get(ProjectInfo.project == wbs)
                wbs_info.append(wbs)
                wbs_info.append('')
                wbs_info.append(result.project_name)
            except ProjectInfo.DoesNotExist:
                return None
            return wbs_info

        try:
            result = ProjectInfo.select(ProjectInfo.project, ProjectInfo.contract, ProjectInfo.project_name, UnitInfo.lift_no, UnitInfo.project_catalog, UnitInfo.nonstd_level,
                                        UnitInfo.is_urgent, UnitInfo.req_configure_finish, UnitInfo.req_delivery_date, ElevatorTypeDefine.elevator_type, SUnitParameter.load, SUnitParameter.speed).join(UnitInfo, on=(UnitInfo.project == ProjectInfo.project))\
                .switch(UnitInfo).join(ElevatorTypeDefine).join(SUnitParameter, on=(SUnitParameter.wbs_no == UnitInfo.wbs_no)).where(UnitInfo.wbs_no == wbs).naive().get()
        except ProjectInfo.DoesNotExist:
            return wbs_info

        wbs_info.append(wbs)  # 0
        wbs_info.append(none2str(result.contract))  # 1
        wbs_info.append(none2str(result.project_name))  # 2
        wbs_info.append(none2str(result.lift_no))  # 3
        wbs_info.append(none2str(result.elevator_type))  # 4
        wbs_info.append(none2str(result.load))  # 5
        wbs_info.append(none2str(result.speed))  # 6
        wbs_info.append(Catalog_Types[result.project_catalog])  # 7
        wbs_info.append(Nonstd_Level[result.nonstd_level])  # 8
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
            b_show = True
            for case in cases:
                try:
                    i = cols.index(case)
                except ValueError:
                    continue

                if self.mat_res[key][i].find(cases[case]) == -1 and self.mat_res[key][i] != cases[case] and cases[case].upper() != 'ALL':
                    b_show = False
                    break

            if not b_show:
                continue

            parent = self.mat_list.insert('', END, values=self.mat_res[key])
            i = cols.index('col1')
            wbs_list = self.nstd_res[self.mat_res[key][i]]

            for wbs in wbs_list:
                self.mat_list.insert(parent, END, values=self.wbs_res[wbs])

    def __refresh_tree(self):
        self.data_thread = refresh_thread(self)
        self.data_thread.setDaemon(True)
        self.data_thread.start()
        # threads.append(data_thread)

    def refresh_by_hand(self):
        if self.data_thread.is_alive():
            messagebox.showinfo('提示', '表单刷新线程正在后台刷新列表，请等待完成后再点击!')
            return

        self.b_finished = False
        self.__refresh_tree()

    def __loop_refresh(self):
        if not self.data_thread:
            pass
        elif self.data_thread.is_alive():
            messagebox.showinfo('提示', '表单刷新线程正在后台运行, 15分钟后自动刷新进程重新启动!')
            time.sleep(900)

        self.b_finished = False
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
                self.mat_entry.delete(i, i + 1)

    def clear_select(self, event):
        items = self.mat_list.selection()
        for item in items:
            self.mat_list.selection_remove(item)

    def copy_mat_list(self, event):
        items = self.mat_list.selection()
        if not items:
            return
        mat_str = ''
        self.mat_list.clipboard_clear()
        for item in items:
            if self.mat_list.parent(item) != '':
                continue
            i = cols.index('col3')
            mat_str = self.mat_list.item(item, 'values')[i] + '\n'
            self.mat_list.clipboard_append(mat_str)

    def copy_list(self, event):
        items = self.mat_list.selection()
        if not items:
            return

        self.mat_list.clipboard_clear()
        mat_str = ''

        for col in display_col:
            i = cols.index(col)
            mat_str = mat_str + tree_head[i] + '\t'
        mat_str = mat_str + '\n'
        self.mat_list.clipboard_append(mat_str)

        for item in items:
            if self.mat_list.parent(item) != '':
                continue
            mat_str = ''
            for col in display_col:
                j = cols.index(col)
                mat_str = mat_str + \
                    self.mat_list.item(item, 'values')[j] + '\t'
            mat_str = mat_str + '\n'
            self.mat_list.clipboard_append(mat_str)

    def select_all(self, event):
        items = self.mat_list.get_children()
        if not items:
            return

        self.mat_list.selection_set(items[0])
        self.mat_list.focus_set()
        self.mat_list.focus(items[0])
        # self.mat_list.selection_remove(items[0])
        for item in items:
            self.mat_list.selection_add(item)
    '''
    def copy_clip_process(self):
        wbs_sel = WBSSelDialog(self, 'WBS List粘贴')
    '''

    def __mat_list(self, choice, mats, comment=None):
        s_error = ''
        i_error = 0
        i_suss = 0
        for mat in mats:
            result = self.process_mat_status(mat.rstrip(), choice, True)
            if result > 0:
                i_suss = i_suss + 1
            else:
                i_error = i_error + 1
                s_error = s_error + mat.rstrip() + ';'

        messagebox.showinfo('结果', str(i_suss) + '更新成功;' +
                            str(i_error) + '更新失败:' + s_error)

        if i_suss >= 1:
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
        if messagebox.askyesno('确认执行', '执行数据数量: ' + str(count) + ' 条;此操作不可逆，是否继续(YES/NO)?') == NO:
            return

        i_suss = 0
        i_error = 0
        s_error = ''
        for item in items:
            if self.mat_list.parent(item) != '':
                continue
            col_index = cols.index('col3')
            mat = self.mat_list.item(item, 'values')[col_index]
            if self.process_mat_status(mat, choice, True) > 0:
                i_suss = i_suss + 1
            else:
                i_error = i_error + 1
                s_error = s_error + mat + ';'

        messagebox.showinfo('结果', str(i_suss) + '更新成功;' +
                            str(i_error) + '更新失败:' + s_error)

        if i_suss >= 1:
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
            query_res = nstd_mat_fin.get(nstd_mat_fin.mat_no == mat)
        except nstd_mat_fin.DoesNotExist:
            return -1

        mat_catalog = query_res.justify
        mat_mbom_fin = query_res.mbom_fin
        mat_pu_price_fin = query_res.pu_price_fin
        mat_co_run_fin = query_res.co_run_fin

        user_per = int(login_info['perm'][1])

        if method == 1:
            # 分类和权限组合判断
            if (mat_catalog == 3 and (user_per == 2 or user_per == 9)) or \
               (mat_catalog == 4 and (user_per == 3 or user_per == 9)) or \
               (mat_catalog == 5 and (user_per == 6 or user_per == 9)):
                if mat_mbom_fin == value:
                    return 2
                else:
                    s = nstd_mat_fin.update(mbom_fin=value, mbom_fin_on=datetime.datetime.now(), mbom_fin_by=login_info[
                                            'uid'], mbom_fin_remark=none2str(comment)).where(nstd_mat_fin.mat_no == mat)
                    i_u = s.execute()
            else:
                return -2
        elif method == 2:
            if (mat_catalog == 1 or mat_catalog == 2 or mat_catalog == 6) and (user_per == 4 or user_per == 9):
                if mat_pu_price_fin == value:
                    return 2
                else:
                    s = nstd_mat_fin.update(pu_price_fin=value, pu_price_fin_on=datetime.datetime.now(), pu_price_fin_by=login_info[
                                            'uid'], pu_price_fin_remark=none2str(comment)).where(nstd_mat_fin.mat_no == mat)
                    i_u = s.execute()
            else:
                return -2
        elif method == 3:
            if user_per != 5 and user_per != 9:
                return -2

            if mat_catalog <= 0:
                return -3

            if (mat_catalog == 1 or mat_catalog == 2 or mat_catalog == 6) and not mat_pu_price_fin:
                return -3

            if (mat_catalog == 3 or mat_catalog == 4 or mat_catalog == 5) and not mat_mbom_fin:
                return -3

            if mat_co_run_fin == value:
                return 2
            else:
                s = nstd_mat_fin.update(co_run_fin=value, co_run_fin_on=datetime.datetime.now(), co_run_fin_by=login_info[
                                        'uid'], co_run_fin_remark=none2str(comment)).where(nstd_mat_fin.mat_no == mat)
                i_u = s.execute()
        else:
            return method

        return i_u
