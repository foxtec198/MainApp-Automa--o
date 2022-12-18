from openpyxl import load_workbook as lw
from time import strftime as st

dia = int(st('%d'))
planilha = lw("metas.xlsx")
ws = planilha.active

class Metas():
    def __init__(self):
        self.soma = float()
        for i in range(4, 5000):
            diaP = ws[f'A{i}'].value
            valor = ws[f'D{i}'].value
            fl = ws[f'B{i}'].value
            if fl == 'L062':
                if diaP == dia:
                    self.soma += valor
    def grupo(self):
        for i in range(4, 5000):
            diaP = ws[f'A{i}'].value
            valor = ws[f'D{i}'].value
            fl = ws[f'B{i}'].value
            grupo = ws[f"C{i}"].value
            if diaP == dia:
                if fl == 'L062':
                    if grupo == 'Basket':
                        self.basket = valor
                    if grupo == 'Beleza':
                        self.blz = valor
                    if grupo == 'Calçados':
                        self.cba = valor
                    if grupo == 'Eletrônicos':
                        self.tec = valor
                    if grupo == 'Feminino':
                        self.fem = valor
                    if grupo == 'Infantil':
                        self.inf = valor
                    if grupo == 'Masculino':
                        self.masc = valor
                    if grupo == 'Moda Casa':
                        self.mdc = valor
                    if grupo == 'Óculos e Bijuteria':
                        self.biju = valor
                    if grupo == 'Relógios':
                        self.rel = valor