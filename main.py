from datetime import date
from EXCEL_FUNCS import *
from DATA_FUNCS import *
from FAEandAP_FUNCS import *

class price_object():
    def __init__(self,path,settings_dic,config_dic):
        self.path_open=path
        self.date_today=''
        if settings_dic["checkBox_date"]:
            self.date_today=str(date.today()).replace("-", "/")
        self.name=settings_dic["name_edit"]
        self.code=settings_dic["code_edit"]
        self.coord_str=settings_dic["coord_edit"]
        self.del_sign=settings_dic["groupBox_null"]
        self.delrow_num=settings_dic["spinBox"]
        self.suffix_sign=settings_dic["groupBox_suffix"]
        self.suffix_content=settings_dic["suffix_edit"]
        self.config_dic=config_dic #{"kvar":[kvar_rbt,checkBox_kvar,checkBox_BKSTSC],"fae":[fire_rbt,checkBox_FAE],"ap":[arc_rbt,checkBox_AP]}

    def operate(self):
        if self.config_dic["kvar"][0]:
            wb=load_workbook(self.path_open)
            ws=wb.active
            row,project_name=ReturnRow(ws)
            region_lst,ws_end_row=TellRegion(ws,row) #[['', 4, 54], ['2号厂房', 56, 80], ['3号厂房', 82, 97], ['4号厂房', 99, 106]]quit()
            data_groups,KVAR_and_A=get_datas(ws,region_lst)
            print(data_groups)
            if self.config_dic["kvar"][2]:#若勾选拆分为BKS+投切开关，对data_groups进行处理
                checkBox_BKSTSC_ischecked(data_groups)
            info_lst=[self.date_today,self.name,self.code]
            wb_price,products_and_rows,bottom_content_start_row,end_row=write_datas(data_groups,region_lst,project_name,info_lst)
            BK_correspondence,products_and_rows_splited=split_BKT_BKC(products_and_rows,self.config_dic["kvar"][1])
            ws_price,product_row_FL,price_row,ws_adj_end_row, coord,BK_correspondence=bulid_form(wb_price,products_and_rows_splited,self.coord_str,BK_correspondence,self.config_dic["kvar"][1])
            set_formula(ws_price,products_and_rows,product_row_FL,price_row,BK_correspondence,ws_adj_end_row, coord,self.config_dic["kvar"][1])
            set_format(ws_price, coord,ws_adj_end_row)# 设置调价表格式
            statistical_table(ws_price, ws_adj_end_row, coord, KVAR_and_A)
            ws_price.print_area = f'A1:H{end_row}' #依据模板来校正打印范围，后续若delete_rows会自动减少对应范围
            try:
                file_save_path=save_file(self.path_open, self.code, project_name, wb_price,"",self.suffix_sign,self.suffix_content)
                if self.del_sign:
                    delete_rows(file_save_path, bottom_content_start_row+self.delrow_num, 200)
                return True
            except:
                return False
        elif self.config_dic["fae"][0]:
            wb = load_workbook(self.path_open)
            ws = wb.active
            start_coord,end_coord=get_the_range(ws)
            nm_r,nm_c=find_the_projectname(ws,start_coord,end_coord)
            project_name=ws.cell(nm_r,nm_c).value[5:]
            coord_lst=find_the_text(ws,start_coord,end_coord,"型号")
            coord_lst=filtration(ws, coord_lst)
            datas_dic=FAE_AP_get_datas(ws, coord_lst)
            combine_sign=False
            if not self.config_dic["fae"][1]:
                datas_dic=combine(datas_dic)
                combine_sign=True
            wb_p=load_workbook("消防灭火--报价模板.xlsx")
            ws_p=wb_p.active
            price_row_dic,datas_dic,coord,useless=FAE_AP_price_adjustment_list(ws_p, datas_dic, self.coord_str,None)
            FAE_AP_write_price(ws_p, price_row_dic, coord,"FAE")
            bottom_content_start_row=fill_in_the_form(ws_p,datas_dic,coord,"FAE",None,combine_sign,None,None)
            start_coord_p, end_coord_p=get_the_range(ws_p)
            date_and_name(ws_p, self.date_today, self.name, project_name,start_coord_p, end_coord_p)
            ws_p.print_area = f'A1:G{coordinate_from_string(ws_p.dimensions.split(":")[1])[1]}'
            try:
                file_save_path=save_file(self.path_open, self.code, project_name, wb_p,"灭火",self.suffix_sign,self.suffix_content)
                if self.del_sign:
                    delete_rows(file_save_path, bottom_content_start_row+self.delrow_num, 200)
                return True
            except:
                return False
        elif self.config_dic["ap"][0]:
            wb = load_workbook(self.path_open)
            ws = wb.active
            start_coord,end_coord=get_the_range(ws)
            nm_r,nm_c=find_the_projectname(ws,start_coord,end_coord)
            project_name=ws.cell(nm_r,nm_c).value[5:]
            coord_lst=find_the_text(ws,start_coord,end_coord,"型号")
            coord_lst=filtration(ws, coord_lst)
            datas_dic=FAE_AP_get_datas(ws, coord_lst)
            combine_sign=False
            if not self.config_dic["ap"][1]:
                datas_dic=combine(datas_dic)
                combine_sign=True
            wb_p=load_workbook("弧光保护--报价模板.xlsx")
            ws_p=wb_p.active
            price_row_dic, datas_dic, coord,price_row_fucai = FAE_AP_price_adjustment_list(ws_p, datas_dic, self.coord_str,"AP")
            FAE_AP_write_price(ws_p, price_row_dic, coord, "AP")
            start_coord_p, end_coord_p=get_the_range(ws_p)
            date_and_name(ws_p, self.date_today, self.name, project_name,start_coord_p, end_coord_p)
            bottom_content_start_row=fill_in_the_form(ws_p,datas_dic,coord,"AP",price_row_fucai,combine_sign, start_coord_p, end_coord_p)
            ws_p.print_area = f'A1:G{coordinate_from_string(ws_p.dimensions.split(":")[1])[1]}'
            try:
                file_save_path = save_file(self.path_open, self.code, project_name, wb_p, "弧光", self.suffix_sign,self.suffix_content)
                if self.del_sign:
                    delete_rows(file_save_path, bottom_content_start_row + self.delrow_num, 350)
                return True
            except:
                return False