import csv
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Dict, List, Optional, Set

from shipping_instruction.config import PMSFileColumnsConfig
from shipping_instruction.util import _get_first_file_in_dir


@dataclass
class PMSRow:
    kata: str
    hin: str
    shipmentDate: Optional[date]
    shipmentQty: int
    shipmentWarehouse: str


@dataclass
class PMSRowsOfKata:
    kata: str
    shipmentDate: date
    shipmentWarehouse: str
    pmsRows: List[PMSRow]

    # def __init__(self,
    #              kata: str,
    #              shipmentDate: date,
    #              shipmentWarehouse: str,
    #              pmsRows: List[PMSRow]):
    #     self.kata = kata
    #     self.shipmentDate = shipmentDate
    #     self.shipmentWarehouse = shipmentWarehouse
    #     self.pmsRows = pmsRows

    @property
    def hins(self) -> List[str]:
        hins: Set[str] = set()
        for pms_row in self.pmsRows:
            hin = pms_row.hin
            hins.add(hin)
        return list(hins)

    @property
    def shipmentQtyOfHin(self) -> Dict[str, int]:
        shipment_qty_of_hin: Dict[str, int] = {}
        for hin in self.hins:
            shipment_qty = 0
            for pms_row in self.pmsRows:
                if hin == pms_row.hin:
                    shipment_qty += pms_row.shipmentQty
            shipment_qty_of_hin[hin] = shipment_qty
        return shipment_qty_of_hin

    @property
    def shipmentQty(self) -> int:
        shipment_qty_of_kata = 0
        for pms_row in self.pmsRows:
            shipment_qty_of_kata += pms_row.shipmentQty
        return shipment_qty_of_kata


class PMSFile:
    __SUFFIX = ".csv"

    def __init__(self,
                 path: str,
                 config: PMSFileColumnsConfig):
        file = _get_first_file_in_dir(path)
        if file is None:
            raise Exception(f"PMS File Not Found: {path}")

        file_p = Path(file)
        if not file_p.suffix == self.__SUFFIX:
            raise Exception(f"Invalid Suffix: {file_p.suffix}")

        # ファイル名の確認に使用
        self.fileName = file_p.name
        C = self.config = config
        # pms から出力されるファイルのエンコードは shift_jis のよう
        with open(str(file_p), newline="", encoding="shift_jis") as csvfile:
            reader = csv.reader(csvfile)
            self.__csvRows = [row for row in reader]

            self.instructionNumber = self.__csvRows[1][C.INSTRUCTION_NUMBER]

            shipment_date = self.__get_valid_shipment_date()
            if shipment_date is None:
                raise Exception("Invalid Shipment Date")

            shipment_warehouse = self.__get_valid_shipment_warehouse()
            if shipment_warehouse is None:
                raise Exception("Invalid Shipment Warehouse")

            # MRP拠点の判定に用いる
            self.headCharOfShipmentWarehouse = shipment_warehouse[0]

            katas: Set[str] = set()
            for index, row in enumerate(self.__csvRows):
                if index == 0:
                    continue

                kata = row[C.KATA]
                katas.add(kata)
            self.katas = list(katas)

            self.pmsRowsOfKatas: List[PMSRowsOfKata] = []
            for kata in self.katas:
                pms_rows: List[PMSRow] = []
                for index, row in enumerate(self.__csvRows):
                    if index == 0:
                        continue

                    if kata == row[C.KATA]:
                        shipment_qty = int(row[C.SHIPMENT_QTY])
                        if shipment_qty > 0:
                            pms_row = PMSRow(kata=kata,
                                             hin=row[C.HIN],
                                             shipmentDate=shipment_date,
                                             shipmentQty=shipment_qty,
                                             shipmentWarehouse=shipment_warehouse)
                            pms_rows.append(pms_row)
                pms_rows_of_a_kata = PMSRowsOfKata(kata=kata,
                                                   shipmentDate=shipment_date,
                                                   shipmentWarehouse=shipment_warehouse,
                                                   pmsRows=pms_rows)
                if pms_rows_of_a_kata.shipmentQty > 0:
                    self.pmsRowsOfKatas.append(pms_rows_of_a_kata)

            if len(self.pmsRowsOfKatas) == 0:
                raise Exception("No Data In PMS File")

    @property
    def pmsRowsOfKata(self) -> Dict[str, PMSRowsOfKata]:
        pms_rows_of_kata: Dict[str, PMSRowsOfKata] = {}
        for kata in self.katas:
            for pms_rows_of_a_kata in self.pmsRowsOfKatas:
                if kata == pms_rows_of_a_kata.kata:
                    pms_rows_of_kata[kata] = pms_rows_of_a_kata
                    break
        return pms_rows_of_kata

    def __get_valid_shipment_date(self) -> Optional[date]:
        C = self.config
        FORMAT = C.SHIPMENT_DATE_FORMAT_VAL
        DEFAULT_SHIPMENT_DATE = datetime.strptime("1900/01/01", FORMAT).date()
        shipment_date = DEFAULT_SHIPMENT_DATE

        for index, row in enumerate(self.__csvRows):
            if index == 0:
                continue

            if index == 1:
                shipment_date = \
                    datetime.strptime(row[C.SHIPMENT_DATE], FORMAT).date()
                continue

            tmp_date = \
                datetime.strptime(row[C.SHIPMENT_DATE], FORMAT).date()
            if not shipment_date == tmp_date:
                return None

        if shipment_date == DEFAULT_SHIPMENT_DATE:
            return None

        return shipment_date

    def __get_valid_shipment_warehouse(self) -> Optional[str]:
        C = self.config
        DEFAULT_SHIPMENT_WAREHOUSE = ""
        shipment_warehouse = DEFAULT_SHIPMENT_WAREHOUSE

        for index, row in enumerate(self.__csvRows):
            if index == 0:
                continue

            if index == 1:
                shipment_warehouse = row[C.SHIPMENT_WAREHOUSE]
                continue

            tmp_warehouse = row[C.SHIPMENT_WAREHOUSE]
            if not shipment_warehouse == tmp_warehouse:
                return None

        if shipment_warehouse == DEFAULT_SHIPMENT_WAREHOUSE:
            return None

        return shipment_warehouse
