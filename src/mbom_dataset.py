#!/usr/bin/env python
# coding:utf-8
"""
  Author:   10256603<mikewolf.li@tkeap.com>
  Purpose:
  Created: 2016/3/15
"""

from peewee import *

# 单位对照表
Unit_Types = (
    ('SET', '套'),
    ('PC', '件'),
    ('M', '米'),
    ('MM', '毫米'),
    ('Kg', '千克'),
    ('m/s', '米/秒'),
    ('s', '秒'),
    ('CM', '分米'),
)

Justify_Types = {
    -1: '未激活',
    0: '未分类',
    1: 'CLU外购',
    2: 'PCU外购',
    3: '电气自制',
    4: '钣金自制',
    5: '曳引机自制',
    6: 'F30外购',
    13: '手动关闭',
}

Catalog_Types = {
    1: 'Common Project',
    2: 'High-Speed Project',
    3: 'Special Project',
    4: 'Major Project',
    5: 'Pre_engineering',
    6: 'Lean Project'
}

Status_Types = {
    -3: "Revised",
    -2: "DEL",
    -1: "CANCEL",
    0: "CRTD",
    1: "ACTIVE",
    2: "FININISH",
    3: "RESTART",
    4: "FREEZED",
    5: "CLOSED",
    6: "RELEASED",
    7: "PART FINISH",
    8: "DELIVERIED",
    9: "PRE-PRODUCTION"
}

Nonstd_Level = {
    1: 'Target STD',
    2: 'Option STD',
    3: 'Simple Non-STD',
    4: 'Complex Non-STD',
    5: 'Comp-Standard',
    6: 'Comp-Measurement',
    7: 'Comp-Configurable',
    11: 'Design Fault-Qty.',
    12: 'Design Fault-Spec.',
    13: 'Sales Order Fault-Qt',
    14: 'Sales Order Fault-Sp',
    15: 'Matl Pick Fault-Qty.',
    16: 'Matl Pick Fault-Spec',
    17: 'Packing Fault-Qty.',
    18: 'Packing Fault-Spec.',
    19: 'Logistic Fault.',
    21: 'Abnormal in logistic',
    22: 'ATI or ECR.',
    23: 'Others'
}

#mbom_db = PostgresqlDatabase('nstd_mat_db_test',user='postgres',password='1q2w3e4r',host='10.127.144.62',)

#mbom_db = PostgresqlDatabase('nstd_mat_db_dev',user='postgres',password='1q2w3e4r',host='localhost',)
mbom_db = PostgresqlDatabase(
               'nstd_mat_db', user='postgres', password='1q2w3e4r', host='10.127.144.62',)


class BaseModel(Model):

    class Meta:
        database = mbom_db


class s_employee(BaseModel):
    employee = CharField(db_column='employee_id',
                         primary_key=True, max_length=8)
    name = CharField(null=True, max_length=32)
    department = CharField(max_length=64, null=True)
    email = CharField(max_length=128, null=True)
    skype_id = CharField(max_length=128, null=True)

    class Meta:
        db_table = 's_employee'


class op_permission(BaseModel):
    employee = ForeignKeyField(s_employee, to_field='employee')
    perm_code = CharField(max_length=64)

    class Meta:
        db_table = 'op_permission'


class login_log(BaseModel):
    employee = CharField(db_column='employee_id', max_length=8)
    log_status = BooleanField(null=True)
    login_time = DateTimeField(null=True)
    logout_time = DateTimeField(null=True)

    class Meta:
        db_table = 'login_log'


class operate_point(BaseModel): 
    employee = ForeignKeyField(s_employee)
    operate_point = CharField(null=True, max_length=64)
    latest_on = DateTimeField(null=True, formats='%Y-%m-%d %H:%M:%S')

    class Meta:
        db_table = 'operate_point'


class s_change_log(BaseModel):
    table_name = CharField(max_length=64, null=False)
    change_section = CharField(max_length=64, null=False)
    key_word = CharField(max_length=64, null=False)
    old_value = TextField(null=False)
    new_value = TextField(null=False)
    log_on = DateTimeField(null=True)
    log_by = CharField(null=True, max_length=64)

    class Meta:
        db_table = 's_change_log'


class nstd_app_head(BaseModel):
    nstd_app = CharField(
        primary_key=True, db_column='nstd_app_id', max_length=16)
    index_mat = CharField(db_column='index_mat_id', max_length=24)
    project = CharField(null=True, db_column='project_id', max_length=18)
    contract = CharField(null=True, db_column='contract_id', max_length=6)
    app_person = CharField(null=True, max_length=64)
    req_fin_on = DateTimeField(formats='%Y-%m-%d %H:%M:%S', null=True)
    create_by = CharField(null=True, max_length=12)
    create_on = DateTimeField(formats='%Y-%m-%d %H:%M:%S', null=True)

    class Meta:
        db_table = 'nstd_app_head'


class nstd_app_link(BaseModel):
    nstd_app = ForeignKeyField(nstd_app_head, on_delete='CASCADE')
    wbs_no = CharField(null=True, max_length=16)
    mbom_fin = BooleanField(default=False)
    mbom_fin_on = DateTimeField(formats='%Y-%m-%d %H:%M:%S', null=True)
    mbom_fin_by = CharField(null=True, max_length=8)

    class Meta:
        primary_key = CompositeKey('nstd_app', 'wbs_no')
        db_table = 'nstd_app_link'


class id_generator(BaseModel):
    id = PrimaryKeyField()
    func_desc = CharField(null=True, max_length=255)
    step = IntegerField(default=1)
    current = IntegerField()
    id_length = IntegerField(default=24)
    pre_character = CharField(null=True, max_length=6)
    fol_character = CharField(null=True, max_length=12)
    remark = TextField(null=True)

    class Meta:
        db_table = 'id_generator'


class mat_basic_info(BaseModel):
    mat_no = CharField(primary_key=True, max_length=18)
    mat_name_en = CharField(null=True, max_length=255)
    mat_name_cn = CharField(max_length=128)
    gross_weight = CharField(null=True)
    base_unit = CharField(null=True, choices=Unit_Types)
    mat_type = CharField(null=True)
    normt = CharField(null=True)
    old_mat_no = CharField(null=True)
    doc_no = CharField(null=True)
    size_dimen = CharField(null=True)
    parameter = CharField(null=True)

    class Meta:
        db_table = 'mat_basic_info'


class struct_group_code(BaseModel):
    st_code = CharField(primary_key=True, max_length=16)
    st_name_en = CharField(null=True, max_length=255)
    st_name_cn = CharField(null=True, max_length=128)
    m_or_e = CharField(null=True, max_length=4)

    class Meta:
        db_table = 'struct_group_code'


class struct_gc_rel(BaseModel):
    st_code = ForeignKeyField(
        struct_group_code, to_field='st_code', on_delete='CASCADE')
    plant = CharField(max_length=6)
    elevator_type = CharField(null=True, max_length=12)
    rp = CharField(null=True, max_length=6)
    box_code = CharField(null=True, max_length=8)

    class Meta:
        db_table = 'struct_gc_rel'


class mat_extra_info(BaseModel):
    mat_no = ForeignKeyField(
        mat_basic_info, to_field='mat_no', on_delete='CASCADE')
    lgfsb = CharField(max_length=4, null=True)
    lgpro = CharField(max_length=4, null=True)
    mrp_controller = CharField(max_length=3, null=True)
    sbdkz = CharField(max_length=1, null=True)
    prod_scheduler = CharField(max_length=3, null=True)
    sobsk = CharField(max_length=2, null=True)
    proc_type = CharField(max_length=1, null=True)
    spec_proc_type = CharField(max_length=2, null=True)
    matgr = CharField(max_length=20, null=True)
    mat_freight_group = CharField(null=True)
    plant = CharField(max_length=4, null=True)
    price_con_ind = CharField(max_length=1, null=True)
    val_type = CharField(max_length=10, null=True)
    val_area = CharField(max_length=4, null=True)
    val_class = CharField(max_length=4, null=True)
    remarks = TextField(null=True)

    class Meta:
        db_table = 'mat_extra_info'


class mat_info(BaseModel):
    mat_no = CharField(primary_key=True, max_length=18)
    mat_name_en = CharField(null=True, max_length=255)
    mat_name_cn = CharField(max_length=128)
    drawing_no = CharField(null=True, max_length=32)
    mat_material = CharField(null=True, max_length=255)
    mat_material_en = CharField(null=True, max_length=255)
    part_weight = DecimalField(max_digits=10, decimal_places=3, null=True)
    mat_unit = CharField(max_length=16, choices=Unit_Types)
    comments = TextField(null=True)
    rp = CharField(null=True, max_length=6)
    rp_zs = CharField(null=True, max_length=6)
    box_code_sj = CharField(null=True, max_length=8)
    box_code_zs = CharField(null=True, max_length=8)
    is_nonstd = BooleanField(default=False)
    is_packing_box = BooleanField(default=False)
    modify_by = CharField(null=True, max_length=12)
    modify_on = DateTimeField(formats='%Y-%m-%d %H:%M:%S', null=True)

    class Meta:
        db_table = 'mat_info'


class bom_header(BaseModel):
    bom_id = CharField(primary_key=True, unique=True, max_length=32)
    plant = CharField(max_length=8)
    mat_no = ForeignKeyField(mat_info, to_field='mat_no')
    drawing_no = CharField(null=True, max_length=32)
    revision = CharField(null=True, max_length=16)
    is_active = BooleanField(default=True)
    struct_code = CharField(null=True, max_length=32)
    modify_by = CharField(null=True, max_length=12)
    modify_on = DateTimeField(formats='%Y-%m-%d %H:%M:%S', null=True)
    create_by = CharField(null=True, max_length=12)
    create_on = DateTimeField(formats='%Y-%m-%d %H:%M:%S', null=True)

    class Meta:
        db_table = 'bom_header'


class bom_item(BaseModel):
    bom_id = ForeignKeyField(
        bom_header, to_field='bom_id', on_delete='CASCADE')
    index = IntegerField()
    st_no = CharField(max_length=32)
    component = ForeignKeyField(mat_info, to_field='mat_no')
    qty = DecimalField(max_digits=10, decimal_places=2)
    bom_remark = TextField(null=True)
    parent_mat = CharField(null=True, max_length=18)
    modify_by = CharField(null=True, max_length=12)
    modify_on = DateTimeField(formats='%Y-%m-%d %H:%M:%S', null=True)
    create_by = CharField(null=True, max_length=12)
    create_on = DateTimeField(formats='%Y-%m-%d %H:%M:%S', null=True)

    class Meta:
        primary_key = CompositeKey('bom_id', 'st_no')
        db_table = 'bom_item'


class nstd_mat_table(BaseModel):
    mat_no = CharField(primary_key=True, max_length=18)
    mat_name_en = CharField(null=True, max_length=255)
    mat_name_cn = CharField(max_length=128)
    drawing_no = CharField(null=True, max_length=32)
    old_drawing_no = CharField(null=True, max_length=32)
    outline_size = CharField(null=True, max_length=64)
    part_type = CharField(null=True, max_length=32)
    mat_loc = CharField(null=True, max_length=32)
    part_weight = CharField(null=True, max_length=32)
    mat_unit = CharField(max_length=16, choices=Unit_Types)
    comments = TextField(null=True)
    rp = CharField(null=True, max_length=6)
    rp_zs = CharField(null=True, max_length=6)
    group_code = CharField(null=True, max_length=16)
    box_code_sj = CharField(null=True, max_length=8)
    box_code_zs = CharField(null=True, max_length=8)
    in_eo_no = CharField(null=True, max_length=64)
    nstd_app = ForeignKeyField(
        nstd_app_head, to_field='nstd_app', on_delete='CASCADE')
    parent = CharField(null=True, max_length=18)
    mat_app_person = CharField(null=True, max_length=64)
    old_mat_no = CharField(null=True, max_length=64)

    class Meta:
        db_table = 'nstd_mat_table'

class working_date(BaseModel):
    date_desc = DateField(primary_key=True, formats='%Y-%m-%d')
    is_working = BooleanField(null=True, default=False)
    
    class Meta:
        db_table = 'working_date'
        
        
class nstd_mat_fin(BaseModel):
    """
    justify:
      0 - 未分类
      1 - CLU外购
      2 - PCU外购
      3 - 电气自制
      4 - 钣金自制
      5 - 曳引机自制
      6 - F30外购
    """
    mat_no = ForeignKeyField(nstd_mat_table, on_delete='CASCADE', unique=True)
    justify = IntegerField(default=0)
    mbom_fin = BooleanField(null=True, default=False)
    mbom_fin_on = DateTimeField(formats='%Y-%m-%d %H:%M:%S', null=True)
    mbom_fin_by = CharField(null=True, max_length=12)
    mbom_fin_remark = TextField(null=True)
    pu_price_fin = BooleanField(null=True, default=False)
    pu_price_fin_on = DateTimeField(formats='%Y-%m-%d %H:%M:%S', null=True)
    pu_price_fin_by = CharField(null=True, max_length=12)
    pu_price_fin_remark = TextField(null=True)
    co_run_fin = BooleanField(null=True, default=False)
    co_run_fin_on = DateTimeField(formats='%Y-%m-%d %H:%M:%S', null=True)
    co_run_fin_by = CharField(null=True, max_length=12)
    co_run_fin_remark = TextField(null=True)
    modify_by = CharField(null=True, max_length=12)
    modify_on = DateTimeField(formats='%Y-%m-%d %H:%M:%S', null=True)
    req_fin_on = DateTimeField(formats='%Y-%m-%d %H:%M:%S', null=True)
    create_by = CharField(null=True, max_length=12)
    create_on = DateTimeField(formats='%Y-%m-%d %H:%M:%S', null=True)  
    active_by = CharField(null=True, max_length=12)
    active_on = DateTimeField(formats='%Y-%m-%d %H:%M:%S', null=True)  
    j_modify_by = CharField(null=True, max_length=12)
    j_modify_on = DateTimeField(formats='%Y-%m-%d %H:%M:%S', null=True) 

    class Meta:
        db_table = 'nstd_mat_fin'
        
class plist_header(BaseModel):
    version_id = CharField(primary_key=True, max_length=64)
    wbs_no = CharField(max_length=16)
    is_active=BooleanField(null=True, default=True)
    create_by = CharField(null=True, max_length=12)
    create_on = DateTimeField(formats='%Y-%m-%d %H:%M:%S', null=True)
    
    class Meta:
        db_table = 'plist_header'
        
class plist_items(BaseModel):
    version_id = ForeignKeyField(plist_header, on_delete='CASCADE')
    item_no = IntegerField()
    mat_no = CharField(max_length=18)
    mat_name = CharField(null=True, max_length=128)
    wbs_element = CharField(max_length=24)
    is_packing_box = BooleanField(default=False)
    box_id = CharField(max_length=12)
    remarks = TextField(null=True)
    modify_by = CharField(null=True, max_length=12)
    modify_on = DateTimeField(formats='%Y-%m-%d %H:%M:%S', null=True)
    create_by = CharField(null=True, max_length=12)
    create_on = DateTimeField(formats='%Y-%m-%d %H:%M:%S', null=True)
    qty = DecimalField(max_digits=10, decimal_places=2)
    
    class Meta:
        db_table = 'plist_items'

class mboxes_table(BaseModel):
    merge_id = CharField( max_length=64)
    version_id = CharField(max_length=64)
    proj_id = CharField(max_length=12)
    wbs_no = CharField(max_length=16)
    box_id = CharField(max_length=12)
    is_active = BooleanField(null=True, default=True)
    modify_by = CharField(null=True, max_length=12)
    modify_on = DateTimeField(formats='%Y-%m-%d %H:%M:%S', null=True)
    create_by = CharField(null=True, max_length=12)
    create_on = DateTimeField(formats='%Y-%m-%d %H:%M:%S', null=True)
    
    class Meta:
        db_table = 'mboxes_table'
     
class box_mat_logic(BaseModel):
    mat_no = CharField(max_length=18)
    box_id = CharField(max_length=12)
    self_prod = BooleanField(null=True)
    door_type = CharField(null=True, max_length=32)
    elevator_type = CharField(null=True, max_length=32)
    open_type = CharField(null=True, max_length=32)
    jamb_size = CharField(null=True, max_length=32)
    fire_resist = BooleanField(null=True)
    door_height_min = IntegerField(null=True)
    door_height_max = IntegerField(null=True)
    door_width_min = IntegerField(null=True)
    door_width_max = IntegerField(null=True)
    packing_qty_min = IntegerField()
    packing_qty_max = IntegerField()
       
    class Meta:
        db_table = 'box_mat_logic'
         
         
class door_packing_mode(BaseModel):
    door_type = CharField(max_length=18)
    box_id = CharField(max_length=12)
    packing_desc  = CharField(max_length=64,null=True)
    packing_except = CharField(max_length=64,null=True)
    g_qty = IntegerField(null = True)
    is_wooden = BooleanField(default=False)
    
    class Meta:
        db_table = 'door_packing_mode'
        
    
pg_db = PostgresqlDatabase(
    'pgworkflow', user='postgres', password='1q2w3e4r', host='10.127.144.62',)


class PgBaseModel(Model):

    class Meta:
        database = pg_db


class ProjectInfo(PgBaseModel):
    branch = CharField(db_column='branch_id', null=True)
    branch_name = CharField(null=True)
    contract = CharField(db_column='contract_id', null=True)
    create_date = DateTimeField(null=True)
    create_emp = CharField(db_column='create_emp_id', null=True)
    modify_date = DateTimeField(null=True)
    modify_emp = CharField(db_column='modify_emp_id', null=True)
    plant = CharField(null=True)
    project = CharField(db_column='project_id', primary_key=True)
    project_name = TextField()
    res_emp = CharField(db_column='res_emp_id', null=True)

    class Meta:
        db_table = 's_project_info'


class ElevatorTypeDefine(PgBaseModel):
    create_date = DateTimeField(null=True)
    create_emp = CharField(db_column='create_emp_id', null=True)
    elevator_type = CharField()
    elevator_type_id = CharField(primary_key=True)
    modify_date = DateTimeField(null=True)
    modify_emp = CharField(db_column='modify_emp_id', null=True)

    class Meta:
        db_table = 's_elevator_type_define'


class UnitInfo(PgBaseModel):
    can_psn = BooleanField(null=True)
    cancel_times = IntegerField(null=True)
    conf_batch = CharField(db_column='conf_batch_id', null=True)
    conf_valid_end = DateField(null=True)
    create_date = DateTimeField(null=True)
    create_emp = CharField(db_column='create_emp_id', null=True)
    e_nstd = CharField(db_column='e_nstd_id', null=True)
    elevator = ForeignKeyField(db_column='elevator_id', null=True,
                               rel_model=ElevatorTypeDefine, to_field='elevator_type_id')
    has_nonstd_inst_info = BooleanField(null=True)
    is_batch = BooleanField(null=True)
    is_urgent = BooleanField(null=True)
    lift_no = CharField(null=True)
    m_nstd = CharField(db_column='m_nstd_id', null=True)
    modify_date = DateTimeField(null=True)
    modify_emp = CharField(db_column='modify_emp_id', null=True)
    nonstd_level = IntegerField(null=True)
    project_catalog = IntegerField(null=True)
    project = CharField(db_column='project_id', null=True)
    req_configure_finish = DateTimeField(null=True)
    req_delivery_date = DateTimeField(null=True)
    restart_times = IntegerField(null=True)
    review_is_urgent = BooleanField(null=True)
    review_valid_end = DateField(null=True)
    status = IntegerField(null=True)
    unit_doc = CharField(db_column='unit_doc_id', null=True)
    version = IntegerField(db_column='version_id', null=True)
    wbs_no = CharField(primary_key=True)
    wf_status = CharField(null=True)
    old_status = IntegerField(null=True)
    old_wf_status = CharField(null=True)

    class Meta:
        db_table = 's_unit_info'


class NonstdAppHeader(PgBaseModel):
    attach = TextField(null=True)
    contract = CharField(db_column='contract_id', null=True)
    create_date = DateTimeField(null=True)
    create_emp = CharField(db_column='create_emp_id', null=True)
    drawing_req_date = DateTimeField(null=True)
    flow_mask = CharField(null=True)
    link_list = TextField(null=True)
    mat_req_date = DateTimeField(null=True)
    modify_date = DateTimeField(null=True)
    modify_emp = CharField(db_column='modify_emp_id', null=True)
    nonstd_desc = TextField(null=True)
    nonstd = CharField(db_column='nonstd_id', primary_key=True)
    project = CharField(db_column='project_id', null=True)
    status = IntegerField(null=True)

    class Meta:
        db_table = 'l_nonstd_app_header'


class NonstdAppItem(PgBaseModel):
    create_date = DateTimeField(null=True)
    create_emp = CharField(db_column='create_emp_id', null=True)
    has_nonstd_draw = BooleanField(null=True)
    has_nonstd_mat = BooleanField(null=True)
    index = CharField(db_column='index_id', primary_key=True)
    item_catalog = CharField(null=True)
    item_ser = IntegerField(null=True)
    link_list = TextField(null=True)
    modify_date = DateTimeField(null=True)
    modify_emp = CharField(db_column='modify_emp_id', null=True)
    nonstd_catalog = CharField(null=True)
    nonstd_desc = TextField(null=True)
    nonstd = CharField(db_column='nonstd_id', null=True)
    nonstd_value = IntegerField(null=True)
    remarks = TextField(null=True)
    res_person = CharField(null=True)
    status = IntegerField(null=True)
    wf_status = CharField(null=True)

    class Meta:
        db_table = 'l_nonstd_app_item'


class NonstdAppItemInstance(PgBaseModel):
    attach_location = TextField(null=True)
    batch = CharField(db_column='batch_id', null=True)
    create_date = DateTimeField(null=True)
    create_emp = CharField(db_column='create_emp_id', null=True)
    has_nonstd_draw = BooleanField(null=True)
    has_nonstd_mat = BooleanField(null=True)
    index = CharField(db_column='index_id', null=True)
    index_mat = CharField(db_column='index_mat_id', null=True)
    ins_finish_date = DateTimeField(null=True)
    ins_is_return = BooleanField(null=True)
    ins_last_return_date = DateTimeField(null=True)
    ins_return_time = IntegerField(null=True)
    ins_start_date = DateTimeField(null=True)
    instance_nstd_desc = TextField(null=True)
    instance_remarks = TextField(null=True)
    modify_date = DateTimeField(null=True)
    modify_emp = CharField(db_column='modify_emp_id', null=True)
    nstd_mat_app = CharField(db_column='nstd_mat_app_id', null=True)
    res_engineer = CharField(null=True)
    status = IntegerField(null=True)

    class Meta:
        db_table = 'l_nonstd_app_item_instance'


class SEmployee(PgBaseModel):
    active = BooleanField(null=True)
    cost_center = CharField(null=True)
    create_date = DateTimeField(null=True)
    create_emp = CharField(db_column='create_emp_id', null=True)
    department = CharField(db_column='department_id', null=True)
    email = CharField(null=True)
    emp_desc = TextField(null=True)
    emp_pos = CharField(null=True)
    employee = CharField(db_column='employee_id', primary_key=True)
    instant = CharField(db_column='instant_id', null=True)
    mobile_phone = CharField(null=True)
    modify_date = DateTimeField(null=True)
    modify_emp = CharField(db_column='modify_emp_id', null=True)
    name = CharField()
    plant = CharField()
    sex = CharField(null=True)
    sub_phone = CharField(null=True)

    class Meta:
        db_table = 's_employee'


class LProcAct(PgBaseModel):
    action = CharField(db_column='action_id')
    action_lead = IntegerField(null=True)
    action_lead_unit = CharField(null=True)
    action_type = CharField(null=True)
    allow_return = BooleanField(null=True)
    create_date = DateTimeField(null=True)
    create_emp = CharField(db_column='create_emp_id', null=True)
    finish_date = DateTimeField(null=True)
    flag = CharField(null=True)
    flow_ser = IntegerField(null=True)
    follow_action = CharField(db_column='follow_action_id', null=True)
    group_catalog = CharField(null=True)
    group = CharField(db_column='group_id', null=True)
    instance = CharField(db_column='instance_id')
    is_active = BooleanField(null=True)
    is_assigned = BooleanField(null=True)
    is_checked = BooleanField(null=True)
    is_end = BooleanField(null=True)
    is_evaluate = BooleanField(null=True)
    is_need_eval = BooleanField(null=True)
    is_restart = BooleanField(null=True)
    is_return = BooleanField(null=True)
    is_start = BooleanField(null=True)
    is_transit = BooleanField(null=True)
    item = BigIntegerField(primary_key=True, db_column='item_id')
    join_mode = IntegerField(null=True)
    modify_date = DateTimeField(null=True)
    modify_emp = CharField(db_column='modify_emp_id', null=True)
    operator = CharField(db_column='operator_id', null=True)
    pre_action = CharField(db_column='pre_action_id', null=True)
    return_time = IntegerField(null=True)
    role = CharField(db_column='role_id', null=True)
    split_mode = IntegerField(null=True)
    start_date = DateTimeField(null=True)
    step_desc = CharField(db_column='step_desc_id', null=True)
    total_flow = IntegerField(null=True)
    urgent_lead = IntegerField(null=True)
    workflow = CharField(db_column='workflow_id', null=True)

    class Meta:
        db_table = 'l_proc_act'


class SUnitParameter(PgBaseModel):
    car_depth = IntegerField(null=True)
    car_door_type = CharField(null=True)
    car_height = IntegerField(null=True)
    car_width = IntegerField(null=True)
    control_sys = CharField(null=True)
    create_date = DateTimeField(null=True)
    create_emp = CharField(db_column='create_emp_id', null=True)
    door_height = IntegerField(null=True)
    door_width = IntegerField(null=True)
    floors = IntegerField(null=True)
    is_through = BooleanField(null=True)
    landing_door_type = CharField(null=True)
    load = IntegerField(null=True)
    modify_date = DateTimeField(null=True)
    modify_emp = CharField(db_column='modify_emp_id', null=True)
    open_type = CharField(null=True)
    para_doc = CharField(db_column='para_doc_id', null=True)
    project = CharField(db_column='project_id', null=True)
    speed = DecimalField(null=True)
    stops = IntegerField(null=True)
    travel_height = IntegerField(null=True)
    wbs_no = CharField(primary_key=True)

    class Meta:
        db_table = 's_unit_parameter'


class WorkflowInfo(PgBaseModel):
    create_date = DateTimeField(null=True)
    create_emp = CharField(db_column='create_emp_id', null=True)
    modify_date = DateTimeField(null=True)
    modify_emp = CharField(db_column='modify_emp_id', null=True)
    parent_workflow = CharField(db_column='parent_workflow_id', null=True)
    workflow_desc = TextField(null=True)
    workflow = CharField(db_column='workflow_id', primary_key=True)
    workflow_name = CharField()
    workflow_status = IntegerField(null=True)

    class Meta:
        db_table = 's_workflow_info'


class ActionInfo(PgBaseModel):
    action_desc = TextField(null=True)
    action = CharField(db_column='action_id', primary_key=True)
    action_name = CharField()
    action_status = IntegerField(null=True)
    create_date = DateTimeField(null=True)
    create_emp = CharField(db_column='create_emp_id', null=True)
    modify_date = DateTimeField(null=True)
    modify_emp = CharField(db_column='modify_emp_id', null=True)

    class Meta:
        db_table = 's_action_info'

class SParameterFields(PgBaseModel):
    field_desc = CharField(null=True)
    field_name = CharField(primary_key=True)

    class Meta:
        db_table = 's_parameter_fields'    
