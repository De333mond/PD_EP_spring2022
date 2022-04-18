from pydantic import BaseModel



class cell(BaseModel):
    modul: int
    discipline: str
    term: str
    zet: float

m = cell(modul=2, discipline="disc", term="first", zet=2.3)

# returns a dictionary:
print(m.dict())