from importlib.util import module_from_spec
from main.main import *
from pydantic import BaseModel;

fullname_db = '.\db.accdb;'
sem = ["Первый", "Второй", "Третий", "Четвертый", "Пятый", "Шестой", "Седьмой", "Восьмой", ]
class cell(BaseModel):
    module: int
    discipline: str
    term: str
    zet: float
    

def main():
    cur = connect_to_DateBase(fullname_db=fullname_db)
    data = select_to_DataBase(cur)
    cell_list = []

    for el in data:
        с = cell(module=el[0],discipline=el[1],term=el[2],zet=el[3])
        cell_list.append(с)

    table = []
    for i in range(0,8):
        temp_list = []

        for el in cell_list:
            if el.term == sem[i] + ' семестр':
                temp_list.append(el.dict())

        table.append(temp_list)

    # print(table)
    for el in table[0]:
        print(el["discipline"])

    return table

if __name__ == "__main__":
     main()