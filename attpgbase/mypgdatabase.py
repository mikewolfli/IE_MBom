from peewee import *

database = PostgresqlDatabase('pgworkflow', **{'host': '10.127.144.62', 'port': 5432, 'password': '1q2w3e4r', 'user': 'postgres'})

class UnknownField(object):
    pass

class BaseModel(Model):
    class Meta:
        database = database

class ClientVersion(BaseModel):
    change_log = TextField(null=True)
    create_on = DateTimeField(null=True)
    is_current = BooleanField(null=True)
    item = BigIntegerField(db_column='item_id', primary_key=True)
    vesion = CharField(db_column='vesion_id', null=True)

    class Meta:
        db_table = 'client_version'

class ExtraNonAppStd(BaseModel):
    box = CharField(db_column='box_id', null=True)
    id = IntegerField()
    is_valid = BooleanField(null=True)
    nstd_catalog = TextField(null=True)
    nstd_part = TextField(null=True)
    nstd_reason = TextField(null=True)
    ratio = DecimalField(null=True)
    rel_part = TextField(null=True)
    type = CharField(null=True)

    class Meta:
        db_table = 'extra_non_app_std'

class LBomChangeInformTable(BaseModel):
    apply_date = DateTimeField(null=True)
    apply_person = CharField(null=True)
    bci_content = TextField(null=True)
    bci = CharField(db_column='bci_id', primary_key=True)
    is_active = BooleanField(null=True)
    is_co = BooleanField(null=True)
    is_cs = BooleanField(null=True)
    is_ie = BooleanField(null=True)
    is_lo = BooleanField(null=True)
    is_major = BooleanField(null=True)
    is_mdm = BooleanField(null=True)
    is_others = BooleanField(null=True)
    is_pm = BooleanField(null=True)
    is_pp = BooleanField(null=True)
    is_psm = BooleanField(null=True)
    is_pu = BooleanField(null=True)
    is_published = BooleanField(null=True)
    is_qm = BooleanField(null=True)
    is_spu = BooleanField(null=True)
    publish_date = DateTimeField(null=True)
    publisher = CharField(null=True)
    remarks = TextField(null=True)

    class Meta:
        db_table = 'l_bom_change_inform_table'

class LCancelLog(BaseModel):
    action = CharField(db_column='action_id', null=True)
    cancel_times = IntegerField(null=True)
    finish_date = DateTimeField(null=True)
    finish_ser = IntegerField(null=True)
    flag = CharField(null=True)
    flow_ser = IntegerField(null=True)
    group = CharField(db_column='group_id', null=True)
    instance = CharField(db_column='instance_id')
    is_finish = BooleanField(null=True)
    item = BigIntegerField(db_column='item_id', primary_key=True)
    operator = CharField(db_column='operator_id', null=True)
    res_desc = CharField(db_column='res_desc_id', null=True)
    start_date = DateTimeField(null=True)
    total_flow = IntegerField(null=True)
    workflow = CharField(db_column='workflow_id')

    class Meta:
        db_table = 'l_cancel_log'

class LChangeLog(BaseModel):
    change_way = CharField(null=True)
    item = BigIntegerField(db_column='item_id')
    key_word = CharField(null=True)
    modify_date = DateTimeField(null=True)
    modify_person = CharField(null=True)
    result_value = TextField(null=True)
    source_char = CharField(null=True)
    source_value = TextField(null=True)
    table_str = CharField(null=True)

    class Meta:
        db_table = 'l_change_log'

class LDailyTask(BaseModel):
    id = BigIntegerField()
    load_unit = CharField(null=True)
    review_load = DecimalField(null=True)
    review_time = DateTimeField(null=True)
    reviewer = CharField(null=True)
    submit_time = DateTimeField(null=True)
    submiter = CharField(null=True)
    task_catalog = CharField(null=True)
    task_content = TextField(null=True)
    task = CharField(db_column='task_id', null=True)
    task_load = DecimalField(null=True)
    task_status = IntegerField(null=True)
    work_date = DateField(null=True)
    work_date_to = DateField(null=True)

    class Meta:
        db_table = 'l_daily_task'

class LEvaluationInstance(BaseModel):
    action = CharField(db_column='action_id', null=True)
    check = CharField(db_column='check_id', null=True)
    create_date = DateTimeField(null=True)
    create_emp = CharField(db_column='create_emp_id', null=True)
    error_point = TextField(null=True)
    error_qty = IntegerField(null=True)
    eval_grade = CharField(null=True)
    eval = CharField(db_column='eval_id', null=True)
    eval_remarks = TextField(null=True)
    eval_type = CharField(null=True)
    eval_value = DecimalField(null=True)
    evaluator = CharField(db_column='evaluator_id', null=True)
    is_valid = BooleanField(null=True)
    item = BigIntegerField(db_column='item_id')
    modify_date = DateTimeField(null=True)
    modify_emp = CharField(db_column='modify_emp_id', null=True)
    op_finish_date = DateTimeField(null=True)
    op_start_date = DateTimeField(null=True)
    operator = CharField(db_column='operator_id', null=True)
    task_desc = TextField(null=True)
    task = TextField(db_column='task_id', null=True)
    workflow = CharField(db_column='workflow_id', null=True)

    class Meta:
        db_table = 'l_evaluation_instance'

class LInstanceStatus(BaseModel):
    action = CharField(db_column='action_id', null=True)
    create_date = DateTimeField(null=True)
    create_emp = CharField(db_column='create_emp_id', null=True)
    instance = CharField(db_column='instance_id')
    modify_date = DateTimeField(null=True)
    modify_emp = CharField(db_column='modify_emp_id', null=True)
    status = IntegerField(null=True)
    wf_status = CharField(null=True)
    workflow = CharField(db_column='workflow_id')

    class Meta:
        db_table = 'l_instance_status'
        indexes = (
            (('instance', 'workflow'), True),
        )
        primary_key = CompositeKey('instance', 'workflow')

class LInternalComTable(BaseModel):
    apply_date = DateTimeField(null=True)
    apply_person = CharField(null=True)
    is_active = BooleanField(null=True)
    is_published = BooleanField(null=True)
    ll_content = TextField(null=True)
    ll = CharField(db_column='ll_id', primary_key=True)
    publish_date = DateTimeField(null=True)
    publisher = CharField(null=True)
    remarks = TextField(null=True)

    class Meta:
        db_table = 'l_internal_com_table'

class LLoginLog(BaseModel):
    employee = CharField(db_column='employee_id')
    item = BigIntegerField(db_column='item_id', primary_key=True)
    log_status = BooleanField(null=True)
    login_ip = CharField(null=True)
    login_pc_name = CharField(null=True)
    login_time = DateTimeField(null=True)
    login_user = CharField(null=True)
    logoff_time = DateTimeField(null=True)
    version = CharField(db_column='version_id', null=True)

    class Meta:
        db_table = 'l_login_log'

class LNonstdAppHeader(BaseModel):
    attach = TextField(null=True)
    contract = CharField(db_column='contract_id', null=True)
    create_date = DateTimeField(null=True)
    create_emp = CharField(db_column='create_emp_id', null=True)
    drawing_req_date = DateTimeField(null=True)
    flow_mask = CharField(null=True)
    id = BigIntegerField()
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

class LNonstdAppItem(BaseModel):
    create_date = DateTimeField(null=True)
    create_emp = CharField(db_column='create_emp_id', null=True)
    has_nonstd_draw = BooleanField(null=True)
    has_nonstd_mat = BooleanField(null=True)
    id = BigIntegerField()
    index = CharField(db_column='index_id', primary_key=True)
    item_catalog = CharField(null=True)
    item_ser = IntegerField(null=True)
    link_list = TextField(null=True)
    modify_date = DateTimeField(null=True)
    modify_emp = CharField(db_column='modify_emp_id', null=True)
    nonstd_catalog = CharField(null=True)
    nonstd_desc = TextField(null=True)
    nonstd = ForeignKeyField(db_column='nonstd_id', null=True, rel_model=LNonstdAppHeader, to_field='nonstd')
    nonstd_value = IntegerField(null=True)
    remarks = TextField(null=True)
    res_person = CharField(null=True)
    status = IntegerField(null=True)
    wf_status = CharField(null=True)

    class Meta:
        db_table = 'l_nonstd_app_item'

class LNonstdAppItemInstance(BaseModel):
    attach_location = TextField(null=True)
    batch = CharField(db_column='batch_id', null=True)
    create_date = DateTimeField(null=True)
    create_emp = CharField(db_column='create_emp_id', null=True)
    has_nonstd_draw = BooleanField(null=True)
    has_nonstd_mat = BooleanField(null=True)
    id = BigIntegerField()
    index = ForeignKeyField(db_column='index_id', null=True, rel_model=LNonstdAppItem, to_field='index')
    index_mat = CharField(db_column='index_mat_id', primary_key=True)
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

class LNonstdConfigureMatList(BaseModel):
    batch = CharField(db_column='batch_id')
    box = CharField(db_column='box_id', null=True)
    create_date = DateTimeField(null=True)
    create_emp = CharField(db_column='create_emp_id', null=True)
    index = CharField(db_column='index_id')
    index_mat = CharField(db_column='index_mat_id', null=True)
    is_active = BooleanField(null=True)
    link_list = TextField(null=True)
    mat_comment = TextField(null=True)
    mat = CharField(db_column='mat_id', null=True)
    mat_name_cn = CharField(null=True)
    mat_name_en = CharField(null=True)
    modify_date = DateTimeField(null=True)
    modify_emp = CharField(db_column='modify_emp_id', null=True)
    plant = CharField(null=True)
    qty = CharField(null=True)
    rp = CharField(null=True)
    sequence = CharField()
    stno = IntegerField(null=True)
    unit = CharField(null=True)

    class Meta:
        db_table = 'l_nonstd_configure_mat_list'
        indexes = (
            (('index', 'batch', 'sequence'), True),
        )
        primary_key = CompositeKey('batch', 'index', 'sequence')

class LNonstdMatList(BaseModel):
    batch = CharField(db_column='batch_id', null=True)
    box = CharField(db_column='box_id', null=True)
    create_date = DateTimeField(null=True)
    create_emp = CharField(db_column='create_emp_id', null=True)
    drawing = CharField(db_column='drawing_id', null=True)
    drawing_needed = BooleanField(null=True)
    index = CharField(db_column='index_id')
    index_mat = CharField(db_column='index_mat_id')
    lift_no = TextField(null=True)
    lift_type = CharField(null=True)
    link_list = TextField(null=True)
    mat_catalog = CharField(db_column='mat_catalog_id', null=True)
    mat_comment = TextField(null=True)
    mat = CharField(db_column='mat_id', primary_key=True)
    mat_name_cn = CharField()
    mat_name_en = CharField()
    mat_status = CharField(null=True)
    modify_date = DateTimeField(null=True)
    modify_emp = CharField(db_column='modify_emp_id', null=True)
    nonstd_designer = CharField(null=True)
    nstd_mat_app = CharField(db_column='nstd_mat_app_id', null=True)
    old_drawing = CharField(null=True)
    plant = CharField(null=True)
    rp = CharField(null=True)
    unit = CharField()

    class Meta:
        db_table = 'l_nonstd_mat_list'

class LPrintLog(BaseModel):
    id = BigIntegerField()
    print_content = TextField(null=True)
    printed_at = DateTimeField(null=True)
    printed_by = CharField(null=True)
    printer = TextField(null=True)

    class Meta:
        db_table = 'l_print_log'

class LProcAct(BaseModel):
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
    item = BigIntegerField(db_column='item_id', primary_key=True)
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

class LRestartLog(BaseModel):
    action = CharField(db_column='action_id', null=True)
    finish_date = DateTimeField(null=True)
    finish_ser = IntegerField(null=True)
    flag = CharField(null=True)
    flow_ser = IntegerField(null=True)
    group = CharField(db_column='group_id', null=True)
    instance = CharField(db_column='instance_id')
    is_finish = BooleanField(null=True)
    item = BigIntegerField(db_column='item_id', primary_key=True)
    operator = CharField(db_column='operator_id', null=True)
    res_desc = CharField(db_column='res_desc_id', null=True)
    restart_times = IntegerField(null=True)
    start_date = DateTimeField(null=True)
    total_flow = IntegerField(null=True)
    workflow = CharField(db_column='workflow_id')

    class Meta:
        db_table = 'l_restart_log'

class LReturnLog(BaseModel):
    from_action = CharField(null=True)
    from_person = CharField(null=True)
    from_start_date = DateTimeField(null=True)
    instance = CharField(db_column='instance_id', null=True)
    item = BigIntegerField(db_column='item_id', primary_key=True)
    rep_doc = CharField(db_column='rep_doc_id', null=True)
    return_date = DateTimeField(null=True)
    return_time_log = IntegerField(null=True)
    to_action = CharField(null=True)
    to_finish_date = DateTimeField(null=True)
    to_person = CharField(null=True)

    class Meta:
        db_table = 'l_return_log'

class LSpecGadRevisedInformTable(BaseModel):
    apply_date = DateTimeField(null=True)
    apply_person = CharField(null=True)
    asg_content = TextField(null=True)
    asg = CharField(db_column='asg_id', primary_key=True)
    is_active = BooleanField(null=True)
    is_published = BooleanField(null=True)
    link = TextField(db_column='link_id', null=True)
    publish_date = DateTimeField(null=True)
    publisher = CharField(null=True)
    remarks = TextField(null=True)
    res_engineer = TextField(null=True)

    class Meta:
        db_table = 'l_spec_gad_revised_inform_table'

class LTracker(BaseModel):
    create_date = DateTimeField(null=True)
    create_emp = CharField(db_column='create_emp_id', null=True)
    instance = CharField(db_column='instance_id', null=True)
    item = IntegerField(db_column='item_id')
    modify_date = DateTimeField(null=True)
    modify_emp = CharField(db_column='modify_emp_id', null=True)
    tracker = CharField(db_column='tracker_id', null=True)

    class Meta:
        db_table = 'l_tracker'

class SActionInfo(BaseModel):
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

class SBranchInfo(BaseModel):
    branch_abb = CharField(null=True)
    branch = CharField(db_column='branch_id', primary_key=True)
    branch_name_cn = CharField(null=True)
    branch_name_en = CharField(null=True)
    is_active = BooleanField(null=True)
    region = CharField(null=True)
    remarks = TextField(null=True)

    class Meta:
        db_table = 's_branch_info'

class SCmBranchRel(BaseModel):
    branch = CharField(db_column='branch_id', null=True)
    employee = CharField(db_column='employee_id', null=True)
    group = CharField(db_column='group_id', null=True)
    id = IntegerField()
    is_valid = BooleanField(null=True)
    region = CharField(null=True)

    class Meta:
        db_table = 's_cm_branch_rel'

class SProjectInfo(BaseModel):
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

class SContractBookHeader(BaseModel):
    contract_doc = CharField(db_column='contract_doc_id', primary_key=True)
    create_date = DateTimeField(null=True)
    create_emp = CharField(db_column='create_emp_id', null=True)
    item_no = IntegerField(null=True)
    modify_date = DateTimeField(null=True)
    modify_emp = CharField(db_column='modify_emp_id', null=True)
    project_catalog = IntegerField(null=True)
    project = ForeignKeyField(db_column='project_id', null=True, rel_model=SProjectInfo, to_field='project')
    status = IntegerField(null=True)

    class Meta:
        db_table = 's_contract_book_header'

class SContractBookInclude(BaseModel):
    contract_doc = CharField(db_column='contract_doc_id', null=True)
    create_date = DateTimeField(null=True)
    create_emp = CharField(db_column='create_emp_id', null=True)
    id = BigIntegerField()
    insert_batch = CharField(db_column='insert_batch_id', null=True)
    is_del = BooleanField(null=True)
    modify_date = DateTimeField(null=True)
    modify_emp = CharField(db_column='modify_emp_id', null=True)
    wbs_no = CharField(null=True)

    class Meta:
        db_table = 's_contract_book_include'

class SContractBrLog(BaseModel):
    b_date = DateTimeField(null=True)
    b_user = CharField(null=True)
    br_status = BooleanField(null=True)
    contract_doc = CharField(db_column='contract_doc_id', null=True)
    item = BigIntegerField(db_column='item_id', primary_key=True)
    r_date = DateTimeField(null=True)
    r_user = CharField(null=True)
    remarks = TextField(null=True)
    s_user = CharField(null=True)

    class Meta:
        db_table = 's_contract_br_log'

class SDepartment(BaseModel):
    create_date = DateTimeField(null=True)
    create_emp = CharField(db_column='create_emp_id', null=True)
    department = CharField(db_column='department_id', primary_key=True)
    department_name = CharField()
    describ = TextField(null=True)
    leader = CharField(null=True)
    modify_date = DateTimeField(null=True)
    modify_emp = CharField(db_column='modify_emp_id', null=True)
    plant = CharField(null=True)

    class Meta:
        db_table = 's_department'

class SDoc(BaseModel):
    doc_desc = TextField(null=True)
    doc = CharField(db_column='doc_id', null=True)

    class Meta:
        db_table = 's_doc'

class SElevatorTypeDefine(BaseModel):
    create_date = DateTimeField(null=True)
    create_emp = CharField(db_column='create_emp_id', null=True)
    elevator_type = CharField()
    elevator_type_id = CharField(primary_key=True)
    modify_date = DateTimeField(null=True)
    modify_emp = CharField(db_column='modify_emp_id', null=True)

    class Meta:
        db_table = 's_elevator_type_define'

class SEmployee(BaseModel):
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

class SEmployeeCardInfo(BaseModel):
    act_date = DateTimeField(null=True)
    card = CharField(db_column='card_id', null=True)
    employee = CharField(db_column='employee_id', null=True)
    id = BigIntegerField(primary_key=True)
    is_active = BooleanField(null=True)
    remarks = TextField(null=True)

    class Meta:
        db_table = 's_employee_card_info'

class SFlowUnitList(BaseModel):
    elevator_type = CharField(db_column='elevator_type_id', null=True)
    group = CharField(db_column='group_id', null=True)
    item = IntegerField(db_column='item_id')
    mode = CharField(null=True)

    class Meta:
        db_table = 's_flow_unit_list'

class SGroup(BaseModel):
    create_date = DateTimeField(null=True)
    create_emp = CharField(db_column='create_emp_id', null=True)
    duty = TextField(null=True)
    group_catalog = CharField(null=True)
    group = CharField(db_column='group_id', primary_key=True)
    group_name = CharField()
    modify_date = DateTimeField(null=True)
    modify_emp = CharField(db_column='modify_emp_id', null=True)
    plant = CharField(null=True)

    class Meta:
        db_table = 's_group'

class SGroupMember(BaseModel):
    create_date = DateTimeField(null=True)
    create_emp = CharField(db_column='create_emp_id', null=True)
    employee = CharField(db_column='employee_id', primary_key=True)
    group_desc = TextField(null=True)
    group = CharField(db_column='group_id', null=True)
    is_leader = BooleanField(null=True)
    is_ncr = BooleanField(null=True)
    item = IntegerField(db_column='item_id')
    leave_backup = CharField(null=True)
    modify_date = DateTimeField(null=True)
    modify_emp = CharField(db_column='modify_emp_id', null=True)
    no_check = BooleanField(null=True)
    status = BooleanField(null=True)

    class Meta:
        db_table = 's_group_member'

class SHierarchy(BaseModel):
    direct_leader = CharField()
    employee = CharField(db_column='employee_id', primary_key=True)
    is_twig = BooleanField(null=True)

    class Meta:
        db_table = 's_hierarchy'

class SHistoryUnitsData(BaseModel):
    config_res_name = CharField(null=True)
    config_table_finish = DateField(null=True)
    config_table_send = DateField(null=True)
    delivery_date = DateField(null=True)
    generate_bom_by_internal = DateField(null=True)
    mat_import_pdm = DateField(null=True)
    project_distribute = DateField(null=True)
    project_mat_prepare = DateField(null=True)
    project_production_plan = DateField(null=True)
    project_release = DateField(null=True)
    project_start = DateField(null=True)
    remarks = TextField(null=True)
    urgent_finish_date = DateField(null=True)
    wbs_no = CharField(primary_key=True)

    class Meta:
        db_table = 's_history_units_data'

class SLeaveDate(BaseModel):
    employee = CharField(db_column='employee_id')
    item = IntegerField(db_column='item_id')
    leave_date = DateField(null=True)

    class Meta:
        db_table = 's_leave_date'

class SLogin(BaseModel):
    employee = CharField(db_column='employee_id', primary_key=True)
    password = CharField(null=True)
    role = CharField(db_column='role_id', null=True)

    class Meta:
        db_table = 's_login'

class SMatForBasicInfo(BaseModel):
    activity = CharField(null=True)
    create_date = DateTimeField(null=True)
    create_emp = CharField(db_column='create_emp_id', null=True)
    mat_no = CharField(primary_key=True)
    mat_para1 = CharField(null=True)
    mat_weight = DecimalField(null=True)
    modify_date = DateTimeField(null=True)
    modify_emp = CharField(db_column='modify_emp_id', null=True)

    class Meta:
        db_table = 's_mat_for_basic_info'

class SNameInfo(BaseModel):
    item_name = CharField(null=True)
    lang_info = CharField(null=True)
    name = CharField(db_column='name_id', primary_key=True)

    class Meta:
        db_table = 's_name_info'

class SNcrData(BaseModel):
    create_by = CharField(null=True)
    create_date = DateTimeField(null=True)
    fb_content = TextField(null=True)
    fb_date = DateTimeField(null=True)
    fetch_date = DateTimeField(null=True)
    file_index = TextField(null=True)
    from_dep = CharField(null=True)
    id = BigIntegerField(primary_key=True)
    issue_desc = TextField(null=True)
    issue_head = TextField(null=True)
    issue_res = CharField(null=True)
    issue_result = TextField(null=True)
    issue_status = IntegerField(null=True)
    lift_no = TextField(null=True)
    lift_qty = IntegerField(null=True)
    modify_by = CharField(null=True)
    modify_date = DateTimeField(null=True)
    ncr = CharField(db_column='ncr_id')
    publish_by = CharField(null=True)
    publish_date = DateTimeField(null=True)
    rel_wbs = TextField(null=True)
    res_bco = TextField(null=True)
    res_by = CharField(null=True)
    res_group = CharField(null=True)
    res_project = TextField(null=True)

    class Meta:
        db_table = 's_ncr_data'

class SNstdDesignOrder(BaseModel):
    act_publish_date = DateTimeField(null=True)
    apply_date = DateTimeField(null=True)
    apply_person = CharField(null=True)
    components_classification = TextField(null=True)
    file = TextField(db_column='file_id', null=True)
    file_name = TextField(null=True)
    is_active = BooleanField(null=True)
    is_car = BooleanField(null=True)
    is_evo = BooleanField(null=True)
    is_evo1 = BooleanField(null=True)
    is_gl = BooleanField(null=True)
    is_hp61_h = BooleanField(null=True)
    is_hp61_l = BooleanField(null=True)
    is_panorama = BooleanField(null=True)
    is_published = BooleanField(null=True)
    is_synergy = BooleanField(null=True)
    is_te_e = BooleanField(null=True)
    ndo = CharField(db_column='ndo_id', primary_key=True)
    nid_components = TextField(null=True)
    plan_publish_date = DateField(null=True)
    prj = TextField(db_column='prj_id', null=True)
    prj_manager = CharField(null=True)
    project_name = TextField(null=True)
    publisher = CharField(null=True)
    remarks = TextField(null=True)
    res_supervior = CharField(null=True)

    class Meta:
        db_table = 's_nstd_design_order'

class SNumberCounter(BaseModel):
    counter_desc = CharField(null=True)
    counter_step = IntegerField(null=True)
    current_counter = BigIntegerField(null=True)
    fol_char = CharField(null=True)
    id_len = IntegerField(null=True)
    item = PrimaryKeyField(db_column='item_id')
    pre_char = CharField(null=True)
    start_counter = IntegerField(null=True)
    table_para_str = CharField(null=True)
    table_str = CharField(null=True)

    class Meta:
        db_table = 's_number_counter'

class SParameterFields(BaseModel):
    field_desc = CharField(null=True)
    field_name = CharField(primary_key=True)

    class Meta:
        db_table = 's_parameter_fields'

class SPartsCatalog(BaseModel):
    catalog_desc = TextField(null=True)
    catalog_name = CharField(null=True)
    part_catalog = CharField(primary_key=True)

    class Meta:
        db_table = 's_parts_catalog'

class SPartsInfo(BaseModel):
    box = CharField(db_column='box_id', null=True)
    elevator_type = CharField(null=True)
    is_delivery = BooleanField(null=True)
    item = IntegerField(db_column='item_id')
    part_catalog = CharField(null=True)
    part_name_en = CharField(null=True)
    part_name_zh = CharField(null=True)
    part_para_notice = CharField(null=True)
    rp = CharField(null=True)
    top_level = CharField(db_column='top_level_id', null=True)

    class Meta:
        db_table = 's_parts_info'

class SWorkflowInfo(BaseModel):
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

class SProcessInfo(BaseModel):
    action = ForeignKeyField(db_column='action_id', rel_model=SActionInfo, to_field='action')
    action_lead = IntegerField(null=True)
    action_lead_unit = CharField(null=True)
    action_status = IntegerField(null=True)
    action_type = CharField(null=True)
    allow_return = BooleanField(null=True)
    create_date = DateTimeField(null=True)
    create_emp = CharField(db_column='create_emp_id', null=True)
    flow_ser = IntegerField(null=True)
    follow_action = CharField(db_column='follow_action_id', null=True)
    group_catalog = CharField(null=True)
    group = CharField(db_column='group_id', null=True)
    is_assigned = BooleanField(null=True)
    is_checked = BooleanField(null=True)
    is_end = BooleanField(null=True)
    is_evaluate = BooleanField(null=True)
    is_need_eval = BooleanField(null=True)
    is_start = BooleanField(null=True)
    is_transit = BooleanField(null=True)
    item = IntegerField(db_column='item_id')
    join_mode = IntegerField(null=True)
    modify_date = DateTimeField(null=True)
    modify_emp = CharField(db_column='modify_emp_id', null=True)
    operator = CharField(db_column='operator_id', null=True)
    pre_action = CharField(db_column='pre_action_id', null=True)
    role = CharField(db_column='role_id', null=True)
    split_mode = IntegerField(null=True)
    total_flow = IntegerField(null=True)
    urgent_lead = IntegerField(null=True)
    workflow = ForeignKeyField(db_column='workflow_id', rel_model=SWorkflowInfo, to_field='workflow')

    class Meta:
        db_table = 's_process_info'
        indexes = (
            (('workflow', 'action'), True),
        )
        primary_key = CompositeKey('action', 'workflow')

class SReviewCommunication(BaseModel):
    id = BigIntegerField()
    is_gad = BooleanField(null=True)
    is_spec = BooleanField(null=True)
    issue_content = TextField(null=True)
    issue_counter = IntegerField(null=True)
    issue_date = DateTimeField(null=True)
    issue_emp = CharField(db_column='issue_emp_id', null=True)
    issue_status = IntegerField(null=True)
    item_no = IntegerField(null=True)
    return_content = TextField(null=True)
    return_counter = IntegerField(null=True)
    return_date = DateTimeField(null=True)
    return_emp = CharField(db_column='return_emp_id', null=True)
    review_task = CharField(db_column='review_task_id', null=True)

    class Meta:
        db_table = 's_review_communication'

class SReviewInfo(BaseModel):
    active_status = IntegerField(null=True)
    create_date = DateTimeField(null=True)
    create_emp = CharField(db_column='create_emp_id', null=True)
    id = BigIntegerField()
    info_remarks = TextField(null=True)
    is_info_full = BooleanField(null=True)
    is_restart = BooleanField(null=True)
    modify_date = DateTimeField(null=True)
    modify_emp = CharField(db_column='modify_emp_id', null=True)
    require_review_date = DateTimeField(null=True)
    res_cm = CharField(null=True)
    review_drawing_qty = IntegerField(null=True)
    review_engineer = CharField(null=True)
    review_remarks = TextField(null=True)
    review_task = CharField(db_column='review_task_id', primary_key=True)
    urgent_level = IntegerField(null=True)

    class Meta:
        db_table = 's_review_info'

class SReviewUnits(BaseModel):
    approve_date = DateTimeField(null=True)
    approver = CharField(null=True)
    create_date = DateTimeField(null=True)
    create_emp = CharField(db_column='create_emp_id', null=True)
    id = BigIntegerField(primary_key=True)
    is_del = BooleanField(null=True)
    is_latest = BooleanField(null=True)
    is_spec_approve = BooleanField(null=True)
    modify_date = DateTimeField(null=True)
    modify_emp = CharField(db_column='modify_emp_id', null=True)
    price_list_receive = DateField(null=True)
    restart_times = IntegerField(null=True)
    review_task = ForeignKeyField(db_column='review_task_id', null=True, rel_model=SReviewInfo, to_field='review_task')
    unit_old_status = IntegerField(null=True)
    unit_status = IntegerField(null=True)
    unit_wf_status = CharField(null=True)
    urgent_level = IntegerField(null=True)
    wbs_no = CharField(null=True)

    class Meta:
        db_table = 's_review_units'

class SReviewUnitsLog(BaseModel):
    item = BigIntegerField(db_column='item_id')
    method = CharField(null=True)
    operated_by = CharField(null=True)
    operated_date = DateTimeField(null=True)
    review_task = CharField(db_column='review_task_id', null=True)
    wbs_no_new = CharField(null=True)
    wbs_no_old = CharField(null=True)

    class Meta:
        db_table = 's_review_units_log'

class SRole(BaseModel):
    create_date = DateTimeField(null=True)
    create_emp = CharField(db_column='create_emp_id', null=True)
    modify_date = DateTimeField(null=True)
    modify_emp = CharField(db_column='modify_emp_id', null=True)
    role = CharField(db_column='role_id', null=True)
    role_name = CharField(null=True)

    class Meta:
        db_table = 's_role'

class SRoleMember(BaseModel):
    create_date = DateTimeField(null=True)
    create_emp = CharField(db_column='create_emp_id', null=True)
    employee = CharField(db_column='employee_id', null=True)
    is_active = BooleanField(null=True)
    item = BigIntegerField(db_column='item_id')
    modify_date = DateTimeField(null=True)
    modify_emp = CharField(db_column='modify_emp_id', null=True)
    role = CharField(db_column='role_id', null=True)

    class Meta:
        db_table = 's_role_member'

class SSapMapTable(BaseModel):
    item = BigIntegerField(db_column='item_id', primary_key=True)
    map_ext = IntegerField(null=True)
    map_method = CharField(null=True)
    para_status = IntegerField(null=True)
    sap_para = CharField(null=True)
    sap_para_name = CharField(null=True)
    sap_table = CharField(null=True)
    we_para = CharField(null=True)
    we_tabel = CharField(null=True)

    class Meta:
        db_table = 's_sap_map_table'

class SSpecAuthProcess(BaseModel):
    action = CharField(db_column='action_id', null=True)
    employee = CharField(db_column='employee_id', null=True)
    id = BigIntegerField(primary_key=True)
    is_valid = BooleanField(null=True)
    workflow = CharField(db_column='workflow_id', null=True)

    class Meta:
        db_table = 's_spec_auth_process'

class SSycLog(BaseModel):
    from_sys = CharField(null=True)
    item = BigIntegerField(db_column='item_id', primary_key=True)
    key1 = CharField(null=True)
    key1_value = TextField(null=True)
    key2 = CharField(null=True)
    key2_value = TextField(null=True)
    key3 = CharField(null=True)
    key3_value = TextField(null=True)
    operator = CharField(db_column='operator_id', null=True)
    syc_desc = TextField(null=True)
    syc_status = IntegerField(null=True)
    syc_time = DateTimeField(null=True)
    to_table = TextField(null=True)

    class Meta:
        db_table = 's_syc_log'

class SSyncToDin(BaseModel):
    balance_rate = DecimalField(null=True)
    car_weight = IntegerField(null=True)
    conf_balance_block_qty = IntegerField(null=True)
    cwt_block_qty_after_dc = IntegerField(null=True)
    cwt_block_qty_before_dc = IntegerField(null=True)
    cwt_block_qty_mat1 = IntegerField(null=True)
    cwt_block_qty_mat2 = IntegerField(null=True)
    cwt_frame_weight = IntegerField(null=True)
    has_governor = BooleanField(null=True)
    has_spring = BooleanField(null=True)
    id = BigIntegerField(primary_key=True)
    project_name = TextField(null=True)
    reserve_decoration_weight = IntegerField(null=True)
    success_info = IntegerField(null=True)
    sync_operator = CharField(null=True)
    sync_time = DateTimeField(null=True)
    wbs_no = CharField(null=True)

    class Meta:
        db_table = 's_sync_to_din'

class STableInfo(BaseModel):
    create_date = DateTimeField(null=True)
    create_emp = CharField(db_column='create_emp_id', null=True)
    modify_date = DateTimeField(null=True)
    modify_emp = CharField(db_column='modify_emp_id', null=True)
    table_desc = TextField(null=True)
    table_name = CharField(db_column='table_name_id', null=True)
    table_str = CharField(primary_key=True)
    table_type = CharField(null=True)

    class Meta:
        db_table = 's_table_info'

class STableParameter(BaseModel):
    create_date = DateTimeField(null=True)
    create_emp = CharField(db_column='create_emp_id', null=True)
    item = IntegerField(db_column='item_id')
    modify_date = DateTimeField(null=True)
    modify_emp = CharField(db_column='modify_emp_id', null=True)
    para_desc = TextField(null=True)
    para_length = IntegerField(null=True)
    para_name = CharField(db_column='para_name_id', null=True)
    para_prec = IntegerField(null=True)
    para_type = CharField(null=True)
    table_para_str = CharField()
    table_str = CharField()

    class Meta:
        db_table = 's_table_parameter'

class STableWorkflow(BaseModel):
    action = CharField(db_column='action_id', null=True)
    item = IntegerField(db_column='item_id')
    table_str = CharField(null=True)
    workflow = CharField(db_column='workflow_id', null=True)

    class Meta:
        db_table = 's_table_workflow'

class STansitions(BaseModel):
    description = TextField(null=True)
    imput = CharField(null=True)
    item = BigIntegerField(db_column='item_id')
    output = CharField(null=True)
    trans_name = CharField(null=True)
    workflow = CharField(db_column='workflow_id', null=True)

    class Meta:
        db_table = 's_tansitions'

class STasksList(BaseModel):
    create_date = DateTimeField(null=True)
    create_emp = CharField(db_column='create_emp_id', null=True)
    is_active = BooleanField(null=True)
    is_new = BooleanField(null=True)
    modify_date = DateTimeField(null=True)
    modify_emp = CharField(db_column='modify_emp_id', null=True)
    refresh_time = IntegerField(null=True)
    task_desc = TextField(null=True)
    task = CharField(db_column='task_id', null=True)
    work = CharField(db_column='work_id', null=True)

    class Meta:
        db_table = 's_tasks_list'

class STractionSpringInfo(BaseModel):
    catagory = IntegerField(null=True)
    description = CharField(null=True)
    id = BigIntegerField(primary_key=True)
    traction_name = CharField(null=True)

    class Meta:
        db_table = 's_traction_spring_info'

class SUnitInfo(BaseModel):
    can_psn = BooleanField(null=True)
    cancel_times = IntegerField(null=True)
    conf_batch = CharField(db_column='conf_batch_id', null=True)
    conf_valid_end = DateField(null=True)
    create_date = DateTimeField(null=True)
    create_emp = CharField(db_column='create_emp_id', null=True)
    e_nstd = CharField(db_column='e_nstd_id', null=True)
    elevator = ForeignKeyField(db_column='elevator_id', null=True, rel_model=SElevatorTypeDefine, to_field='elevator_type_id')
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

    class Meta:
        db_table = 's_unit_info'

class SUnitInfoAttach(BaseModel):
    balance_rate = DecimalField(null=True)
    buffer_man_no_car = CharField(null=True)
    buffer_man_no_cwt = CharField(null=True)
    car_weight = IntegerField(null=True)
    conf_balance_block0_qty = IntegerField(null=True)
    conf_balance_block_qty = IntegerField(null=True)
    create_date = DateTimeField(null=True)
    create_emp = CharField(db_column='create_emp_id', null=True)
    cwt_block_qty_after_dc = IntegerField(null=True)
    cwt_block_qty_before_dc = IntegerField(null=True)
    cwt_block_qty_mat1 = IntegerField(null=True)
    cwt_block_qty_mat2 = IntegerField(null=True)
    cwt_frame_weight = IntegerField(null=True)
    govern_man_no_car = CharField(null=True)
    govern_man_no_cwt = CharField(null=True)
    has_governor = BooleanField(null=True)
    has_spring = BooleanField(null=True)
    is_basic_finish = BooleanField(null=True)
    is_basic_full_finish = BooleanField(null=True)
    is_extern_finish = BooleanField(null=True)
    is_sync_din_latest = BooleanField(null=True)
    mainboard_man_no = CharField(null=True)
    modify_date = DateTimeField(null=True)
    modify_emp = CharField(db_column='modify_emp_id', null=True)
    reserve_decoration_weight = IntegerField(null=True)
    safty_guard_man_no_car = CharField(null=True)
    safty_guard_man_no_cwt = CharField(null=True)
    tm_spring_length_left = IntegerField(null=True)
    tm_spring_length_right = IntegerField(null=True)
    traction_man_no = CharField(null=True)
    wbs_no = CharField(primary_key=True)

    class Meta:
        db_table = 's_unit_info_attach'

class SUnitInfoAttachPrintLog(BaseModel):
    id = BigIntegerField()
    printed_by = CharField(null=True)
    printed_date = DateTimeField(null=True)
    printer = TextField(null=True)
    wbs_no = CharField(null=True)

    class Meta:
        db_table = 's_unit_info_attach_print_log'

class SUnitNstdLevel(BaseModel):
    create_date = DateTimeField(null=True)
    create_emp = CharField(db_column='create_emp_id', null=True)
    is_sap_update = BooleanField(null=True)
    modify_date = DateTimeField(null=True)
    modify_emp = CharField(db_column='modify_emp_id', null=True)
    nonstd_level_pre = IntegerField(null=True)
    wbs_no = CharField()

    class Meta:
        db_table = 's_unit_nstd_level'

class SUnitParaFields(BaseModel):
    create_date = DateTimeField(null=True)
    create_emp = CharField(db_column='create_emp_id', null=True)
    field_name = CharField(null=True)
    field_value = CharField(null=True)
    id = BigIntegerField(primary_key=True)
    modify_date = DateTimeField(null=True)
    modify_emp = CharField(db_column='modify_emp_id', null=True)
    wbs_no = CharField(null=True)

    class Meta:
        db_table = 's_unit_para_fields'

class SUnitParameter(BaseModel):
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

class SVersionLog(BaseModel):
    change_date = DateTimeField(null=True)
    change_doc = CharField(db_column='change_doc_id', null=True)
    item = BigIntegerField(db_column='item_id')
    version = CharField(db_column='version_id', null=True)
    wbs_no = CharField(null=True)

    class Meta:
        db_table = 's_version_log'

class SWorkflowInstance(BaseModel):
    table_para_str = CharField(null=True)
    table_str = CharField(null=True)
    workflow = CharField(db_column='workflow_id', primary_key=True)

    class Meta:
        db_table = 's_workflow_instance'

