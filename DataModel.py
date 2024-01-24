from docx import Document
from docx.shared import Pt


class DataModel:


    data_for_first_page = {"ProjectName": "",
            "Dev": "",
            "Verify": "",
            "NControl": "",
            "Approve": "",

            "Scheme": True,
            "PE": True,
            "DK": True,
            "I1": True,
            "I2": True,

                           "SB_count": 1,
                           "Precent": 10}
    '''Данные для переноса в рамку на главной странице'''



    TemplatesPath: str = ''
    '''Путь к шаблонам'''



    counter = 1
    dfs = []
    '''DataFrame'ы с изначальными данными'''

    dict_chars = {}
    '''Словарь комбинированных изначальных DataFrame'ов с их буквенными обозначениями'''

    final_df = []
    '''Список финальных DataFrame'ов на вставку в документ(-ы)'''

    names_df = {}
    '''Словарь с формированными строками для вставки в документы'''

    files = []
    '''Список названий файлов для обработки'''

    comment_categories = {}
    '''Категории примечаний'''

    not_install_category = {}
    '''Категории для "Не устанавливать"'''

    need_combine_flag = False

    components_one_manuf_categories = {}
    '''Категории компонентов одного производителя'''

    civilian: bool
