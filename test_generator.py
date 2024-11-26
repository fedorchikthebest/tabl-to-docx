from docxtpl import DocxTemplate
from openpyxl import load_workbook


classes = [
    'работник, назначенный в качестве лица, ответственного за обеспечение транспортной безопасности в субъекте транспортной инфраструктуры',
    'работник, назначенный в качестве лица, ответственного за обеспечение транспортной безопасности на объекте транспортной инфраструктуры и (или) транспортном средстве, и персонала специализированных организаций',
    'работник субъекта транспортной инфраструктуры, подразделения транспортной безопасности, руководящий выполнением работы, непосредственно связанной с обеспечением транспортной безопасности объекта транспортной инфраструктуры и (или) транспортного средства',
    'работник, включенный в состав группы быстрого реагирования',
    'работник, осуществляющий досмотр, дополнительный досмотр и повторный досмотр в целях обеспечения транспортной безопасности',
    'работник, осуществляющий наблюдение и (или) собеседование в целях обеспечения транспортной безопасности',
    'работник, управляющий техническими средствами обеспечения транспортной безопасности'
    'иной работник субъекта транспортной инфраструктуры, подразделения транспортной безопасности, выполняющий работы, непосредственно связанные с обеспечением транспортной безопасности объекта транспортной инфраструктуры и (или) транспортного средства'
]


def render_shablons(csv_path, save_path):
    wb_obj = load_workbook(csv_path)
    sheet_obj = wb_obj.active
    
    doc = DocxTemplate("shablon.docx")
    doc_reshenie = DocxTemplate("attest_shablon.docx")
    row_num = 2
    
    while sheet_obj.cell(row = row_num, column = 1).value is not None:
        contexts = [sheet_obj.cell(row_num, i).value for i in range(1, 12)]
        context = {'name' : contexts[0],
                'born_date' : contexts[1],
                'attes_num' : contexts[2],
                'start_date' : contexts[3],
                'stop_date' : contexts[4],
                'a_class' : classes[int(contexts[5]) - 1],
                'num' : contexts[6],
                'email': contexts[7],
                'company': contexts[8],
                'podpis': contexts[9],
                'reshenie_date': contexts[10]}

        doc.render(context)
        doc.save(f"{save_path}\\{context['name']}_{context['num']}_{row_num}.docx")
        
        doc_reshenie.render(context)
        doc_reshenie.save(f"{save_path}\\reshenie_{context['name']}_{context['num']}_{row_num}.docx")
        
        row_num += 1


render_shablons("./data.xlsx", "./result")
