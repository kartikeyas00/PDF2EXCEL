# -*- coding: utf-8 -*-

import tabula
from PyQt5 import QtCore
from PyPDF2 import PdfFileReader
import pandas as pd
import re

class PdfObject(QtCore.QObject):
    
    progressChanged = QtCore.pyqtSignal(int)
    maximumChanged = QtCore.pyqtSignal(int)
    pandasChanged = QtCore.pyqtSignal(pd.DataFrame)

    @QtCore.pyqtSlot(str)
    def pdf2excel(self, pdf_file):
        """ Converts a certain kind of PDF file to Excel with Tabula module"""
        
        pdf = PdfFileReader(open(pdf_file, "rb"))
        length = pdf.getNumPages()
        result = pd.DataFrame(
            columns=[
                "Department",
                "Employment No",
                "Employment Name",
                "Hire Date",
                "Term Date",
                "Birth Date",
                "Seniority Date",
                "Pay Code",
                "FT/PT/S",
                "Status",
            ]
        )
        self.maximumChanged.emit(length)
        page = 1
        while page <= length:
            self.progressChanged.emit(page)
            df = tabula.read_pdf(
                pdf_file,
                pages=str(page),
                lattice=True,
                area=(75.775, 16.0, 572.715, 779.29),
            )[1:]
            pattern = re.compile(r"(\s){2,}")
            df = pd.DataFrame(df[df.columns[0]].replace(pattern, ","))
            df = df["Unnamed: 0"].str.split(",", expand=True)
            df = df.rename(
                columns={
                    0: "Department",
                    1: "Employment No",
                    2: "Employment Name",
                    3: "Hire Date",
                    4: "Term Date",
                    5: "Birth Date",
                    6: "Seniority Date",
                    7: "Pay Code",
                    8: "FT/PT/S",
                    9: "Status",
                }
            )
            result = result.append(df, ignore_index=True)
            page += 1
        result["Hire Date"] = pd.to_datetime(result["Hire Date"])
        result["Term Date"] = pd.to_datetime(result["Term Date"])
        result["Days Difference"] = (
            result["Term Date"] - result["Hire Date"]
        ).dt.days
        result = result.dropna(how="all")
        result = result.drop(columns=["Birth Date", "Pay Code", "Status"])
        result = result[
            [
                "Department",
                "Employment No",
                "Employment Name",
                "Hire Date",
                "Term Date",
                "Days Difference",
                "Seniority Date",
                "FT/PT/S",
            ]
        ]
        self.pandasChanged.emit(result)