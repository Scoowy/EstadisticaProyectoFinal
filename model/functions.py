#!/usr/bin/python
# -*- coding: utf-8 -*-
# Scoowy - Juan Gahona

import pandas as pd
from pandas import DataFrame

from model.file import PandasExcel, ExcelDoc


def createTableCuantitative():
    cuantitativaDf = PandasExcel(
        'VariableCuantitativa.xlsx', 'Datos').openDoc()

    cuantitativaDf = Operations(cuantitativaDf).mergeDfs()

    # print(cuantitativaDf)

    output = ExcelDoc('VariableCuantitativa-output.xlsx', cuantitativaDf)
    output.saveDoc()

    print('Archivo vreado correctamente en ./data/VariableCuantitativa-output.xlsx')


class Operations(object):

    def __init__(self, df: DataFrame, columns=''):
        self.columns = {}
        self.df = df

    def calculateColumns(self):
        Xi = self.df.Xi.values
        fi = self.df.fi.values
        Fi = []
        hi = []
        porcen = []
        Hi = []

        total = 0
        for number in fi:
            total += number

        aux = 0
        aux2 = 0
        for index in range(len(fi)):
            Fi.append(fi[index] + aux)
            aux = Fi[index]

            hi.append(round(fi[index]/total, 3))
            # porcen.append(round(hi[index]*100, 2))
            porcen.append(hi[index])

            Hi.append(round(hi[index] + aux2, 3))
            aux2 = Hi[index]

        self.columns = {'Xi': Xi, 'Fi': Fi, 'hi': hi, '%': porcen, 'Hi': Hi}

        return DataFrame(self.columns)

    def mergeDfs(self):
        return pd.merge(self.df, self.calculateColumns(), on='Xi')
