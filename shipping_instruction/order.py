from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from typing import ClassVar, Dict, List, Optional, Set

from xlrd import open_workbook, sheet, xldate
from xlwt import Workbook, Worksheet

from shipping_instruction.config import OrderFileColumnConfingBase
from shipping_instruction.pms import PMSFile, PMSRow
from shipping_instruction.util import _init_dir


@dataclass
class Order:
    __DEFAULT_UPDATED_RELEASED_QTY: ClassVar[int] = 0

    orderID: str
    orderNumber: str
    kata: str
    orderQty: int
    isNew: bool
    releasedQty: int
    splRows: List["SPLRow"] = field(default_factory=list)
    releasedRows: List[int] = field(default_factory=list)
    notReleasedRows: List[int] = field(default_factory=list)

    # def __init__(self,
    #              orderID: str,
    #              orderNumber: str,
    #              kata: str,
    #              orderQty: int,
    #              isNew: bool,
    #              releasedQty: int):
    #     self.orderID = orderID
    #     self.orderNumber = orderNumber
    #     self.kata = kata
    #     self.orderQty = orderQty
    #     self.isNew = isNew
    #     self.releasedQty = releasedQty

    def __post_init__(self):
        self.__updatedReleasedQty = self.__DEFAULT_UPDATED_RELEASED_QTY
        # self.splRows: List[SPLRow] = []
        # self.releasedRows: List[int] = []
        # self.notReleasedRows: List[int] = []

    @property
    def updatedReleasedQty(self) -> int:
        return self.__updatedReleasedQty

    @updatedReleasedQty.setter
    def updatedReleasedQty(self, value: int):
        if value < self.__DEFAULT_UPDATED_RELEASED_QTY:
            raise Exception("Updated Released Quantity Less Default")

        if value > self.releasedQty:
            raise Exception("Updated Released Quantity Over Released Quantity")

        self.__updatedReleasedQty = value

    @property
    def isUpdateDone(self) -> bool:
        return self.updatedReleasedQty == self.releasedQty

    @property
    def notUpdatedShipmentQty(self) -> int:
        return self.releasedQty - self.updatedReleasedQty

    @property
    def hasNotTBDSPLRow(self) -> bool:
        return len(self.notTBDSPLRows) >= 1

    @property
    def originalRows(self) -> List[int]:
        original_rows = []
        original_rows.extend(self.releasedRows)
        original_rows.extend(self.notReleasedRows)
        return original_rows

    @property
    def notTBDSPLRows(self) -> List[SPLRow]:
        not_tbd_spl_rows = []
        for spl_row in self.splRows:
            if not spl_row.isTBD:
                not_tbd_spl_rows.append(spl_row)

        return not_tbd_spl_rows


class OrderFile:
    __SUFFIX = ".xls"

    def __init__(self,
                 isNew: bool,
                 path: str,
                 config: OrderFileColumnConfingBase):

        path_p = Path(path)
        if not path_p.is_file():
            raise Exception(f"File Not Exist: {path}")

        if not path_p.suffix == self.__SUFFIX:
            raise Exception(
                f"Expected Suffix {self.__SUFFIX}, Get {path_p.suffix}")

        C = self.OrderColumns = config

        self.isNew = isNew
        workbook = open_workbook(path,
                                 formatting_info=True,
                                 on_demand=True)
        self.worksheet: sheet.Sheet = workbook.sheet_by_name(C.SHEET)
        sh = self.worksheet

        # オーダごとのリリース数量を調べるために ID が必要
        order_ids: Set[str] = set()
        for row in range(1, sh.nrows):
            order_ids.add(
                str(int(sh.cell(row, C.JUTYUU_ID).value))
            )

        self.orders: List[Order] = []
        for order_id in order_ids:
            for row in range(1, sh.nrows):
                tmp_order_id = str(int(sh.cell(row, C.JUTYUU_ID).value))
                if order_id == tmp_order_id:
                    order_qty = int(sh.cell(row, C.JUTYUU_SUU).value)
                    default_released_qty = order_qty if self.isNew else 0
                    order = Order(orderID=order_id,
                                  orderNumber=str(
                                      int(sh.cell(row, C.JUTYUU_ORDER_BANGOU).value)
                                  ),
                                  kata=str(sh.cell(row, C.KATABAN).value),
                                  orderQty=order_qty,
                                  isNew=self.isNew,
                                  releasedQty=default_released_qty)
                    self.orders.append(order)
                    break

        if not self.isNew:
            for order in self.orders:
                for row in range(1, sh.nrows):
                    tmp_order_id = str(int(sh.cell(row, C.JUTYUU_ID).value))
                    if tmp_order_id == order.orderID:
                        tmp_status = str(sh.cell(row, C.SYUKKA_STATUS).value)
                        if tmp_status == C.RELEASED_VAL:
                            released_qty = int(
                                sh.cell(row, C.KAITOU_SUU).value
                            )
                            order.releasedQty += released_qty
                            order.releasedRows.append(row)
                        else:
                            order.notReleasedRows.append(row)

        if len(self.orders) == 0:
            raise Exception("No Data In Order File")

    @property
    def katas(self) -> List[str]:
        katas: Set[str] = set()
        for order in self.orders:
            katas.add(order.kata)
        return list(katas)

    @property
    def ordersOfKata(self) -> Dict[str, List[Order]]:
        orders_of_kata: Dict[str, List[Order]] = {}
        for kata in self.katas:
            orders: List[Order] = []
            for order in self.orders:
                if kata == order.kata:
                    orders.append(order)
            orders_of_kata[kata] = orders
        return orders_of_kata

    @property
    def releasedQtyOfKata(self) -> Dict[str, int]:
        released_qty_of_kata: Dict[str, int] = {}
        for kata, orders in self.ordersOfKata.items():
            released_qty = 0
            for order in orders:
                released_qty += order.releasedQty
            released_qty_of_kata[kata] = released_qty
        return released_qty_of_kata

    @property
    def ordersHasNotTBDSPLRow(self) -> List[Order]:
        orders_has_not_tbd_rows: List[Order] = list()
        for order in self.orders:
            if order.hasNotTBDSPLRow:
                orders_has_not_tbd_rows.append(order)
        return orders_has_not_tbd_rows

    def output_upload_file(self, output: str) -> bool:
        output_p = Path(output)
        if output_p.is_dir():
            raise Exception(f"[{output}] is directory")

        if not output_p.suffix == self.__SUFFIX:
            raise Exception(
                f"Output Suffix Not [{self.__SUFFIX}]: {output_p.suffix}")

        if _init_dir(str(output_p.parent), False) is None:
            raise Exception(
                f"Directory Initialize Error: {str(output_p.parent)}"
            )

        C = self.OrderColumns

        rd_sh = self.worksheet
        wt_wb: Workbook = Workbook()
        wt_sh: Worksheet = wt_wb.add_sheet(C.SHEET)

        wt_row = 1
        for order in self.ordersHasNotTBDSPLRow:
            for org_row in order.originalRows:
                wt_sh.write(wt_row, C.JUTYUU_ID,
                            order.orderID)

                self.__copy_column(sheetFrom=rd_sh, sheetTo=wt_sh,
                                   rowFrom=org_row, rowTo=wt_row,
                                   col=C.JUTYUU_RECORD_KOUSHIN_BI,
                                   asDatetime=True)

                self.__copy_column(sheetFrom=rd_sh, sheetTo=wt_sh,
                                   rowFrom=org_row, rowTo=wt_row,
                                   col=C.NOUKI_KAITOU_HDR_RECORD_KOUSHIN_BI,
                                   asDatetime=True)

                self.__copy_column(sheetFrom=rd_sh, sheetTo=wt_sh,
                                   rowFrom=org_row, rowTo=wt_row,
                                   col=C.NOUKI_KAITOU_DTL_RECORD_KOUSHIN_BI,
                                   asDatetime=True)

                self.__copy_column(sheetFrom=rd_sh, sheetTo=wt_sh,
                                   rowFrom=org_row, rowTo=wt_row,
                                   col=C.NOUKI_KAITOU_DID)

                self.__copy_column(sheetFrom=rd_sh, sheetTo=wt_sh,
                                   rowFrom=org_row, rowTo=wt_row,
                                   col=C.HINBAN)

                self.__copy_column(sheetFrom=rd_sh, sheetTo=wt_sh,
                                   rowFrom=org_row, rowTo=wt_row,
                                   col=C.KAITOU_SUU)

                self.__copy_column(sheetFrom=rd_sh, sheetTo=wt_sh,
                                   rowFrom=org_row, rowTo=wt_row,
                                   col=C.KAITOU_SYUKKA_BI,
                                   asDate=True)

                self.__copy_column(sheetFrom=rd_sh, sheetTo=wt_sh,
                                   rowFrom=org_row, rowTo=wt_row,
                                   col=C.MOKUHYOU_NOUKI,
                                   asDate=True)

                self.__copy_column(sheetFrom=rd_sh, sheetTo=wt_sh,
                                   rowFrom=org_row, rowTo=wt_row,
                                   col=C.HIKARI_MRP_KAITOU_NOUKI,
                                   asDate=True)

                self.__copy_column(sheetFrom=rd_sh, sheetTo=wt_sh,
                                   rowFrom=org_row, rowTo=wt_row,
                                   col=C.SPEC_TYOKUSOU)

                self.__copy_column(sheetFrom=rd_sh, sheetTo=wt_sh,
                                   rowFrom=org_row, rowTo=wt_row,
                                   col=C.NAMAMUGI)

                self.__copy_column(sheetFrom=rd_sh, sheetTo=wt_sh,
                                   rowFrom=org_row, rowTo=wt_row,
                                   col=C.SYUKKA_SOUKO)

                self.__copy_column(sheetFrom=rd_sh, sheetTo=wt_sh,
                                   rowFrom=org_row, rowTo=wt_row,
                                   col=C.YOTAKUSAKI_SOUKO)

                self.__copy_column(sheetFrom=rd_sh, sheetTo=wt_sh,
                                   rowFrom=org_row, rowTo=wt_row,
                                   col=C.TEISEI_RIYUU_C)

                self.__copy_column(sheetFrom=rd_sh, sheetTo=wt_sh,
                                   rowFrom=org_row, rowTo=wt_row,
                                   col=C.TYUUSYAKU)

                if org_row in order.releasedRows:
                    wt_sh.write(wt_row, C.SAKUJO_F,
                                C.DELETE_VAL)
                    wt_sh.write(wt_row, C.TEISEI_RIYUU_SYOSAI_NAIYOU,
                                C.TEISEKI_RIYUU_VAL)
                else:
                    self.__copy_column(sheetFrom=rd_sh, sheetTo=wt_sh,
                                       rowFrom=org_row, rowTo=wt_row,
                                       col=C.TEISEI_RIYUU_SYOSAI_NAIYOU)

                wt_row += 1

            for spl_row in order.splRows:
                wt_sh.write(wt_row, C.JUTYUU_ID,
                            order.orderID)

                wt_sh.write(wt_row, C.HINBAN,
                            spl_row.hin)

                wt_sh.write(wt_row, C.KAITOU_SUU,
                            spl_row.shipmentQty)

                if spl_row.shipmentDate is None:
                    wt_sh.write(wt_row, C.KAITOU_SYUKKA_BI,
                                "")
                    wt_sh.write(wt_row, C.MOKUHYOU_NOUKI,
                                "")
                else:
                    wt_sh.write(wt_row, C.KAITOU_SYUKKA_BI,
                                str(spl_row.shipmentDate))
                    wt_sh.write(wt_row, C.MOKUHYOU_NOUKI,
                                str(spl_row.shipmentDate))

                wt_sh.write(wt_row, C.SYUKKA_SOUKO,
                            spl_row.shipmentWarehouse)

                if not self.isNew:
                    wt_sh.write(wt_row, C.TEISEI_RIYUU_SYOSAI_NAIYOU,
                                C.TEISEKI_RIYUU_VAL)

                wt_sh.write(wt_row, C.TYUUSYAKU,
                            C.TYUUSYAKU_VAL)

                wt_row += 1

        if wt_row >= 2:
            wt_wb.save(output)
            return output_p.is_file()
        else:
            return False

    @classmethod
    def __copy_column(cls,
                      sheetFrom: sheet.Sheet,
                      sheetTo: Worksheet,
                      rowFrom: int,
                      rowTo: int,
                      col: Optional[int],
                      asDate: bool = False,
                      asDatetime: bool = False) -> str:

        def __wt(value: str):
            sheetTo.write(rowTo, col,
                          value)

        value = sheetFrom.cell(rowFrom, col).value
        if value == "":
            __wt("")
            return ""

        if asDate:
            newValue = str(xldate.xldate_as_datetime(
                value, sheetFrom.book.datemode).date())
            __wt(newValue)
            return newValue

        if asDatetime:
            newValue = str(xldate.xldate_as_datetime(
                value, sheetFrom.book.datemode))
            __wt(newValue)
            return newValue

        try:
            newValue = str(int(value))
        except ValueError:
            newValue = str(value)
        __wt(newValue)
        return newValue


@dataclass
class SPLRow(PMSRow):
    __DEFAULT_COPIED_SHIPMENT_QTY: ClassVar[int] = 0

    # kata: str
    # hin: str
    # shipmentDate: Optional[date]
    # shipmentQty: int
    # shipmentWarehouse: str
    isTBD: bool

    # def __init__(self,
    #              kata: str,
    #              hin: str,
    #              shipmentDate: Optional[date],
    #              shipmentQty: int,
    #              shipmentWarehouse: str,
    #              isTBD: bool):

    #     super().__init__(kata=kata,
    #                      hin=hin,
    #                      shipmentDate=shipmentDate,
    #                      shipmentQty=shipmentQty,
    #                      shipmentWarehouse=shipmentWarehouse)

    #     self.__copiedShipmentQty = self.__DEFAULT_COPIED_SHIPMENT_QTY
    #     self.isTBD = isTBD

    def __post_init__(self):
        self.__copiedShipmentQty = self.__DEFAULT_COPIED_SHIPMENT_QTY

    @property
    def copiedShipmentQty(self) -> int:
        return self.__copiedShipmentQty

    @copiedShipmentQty.setter
    def copiedShipmentQty(self, value: int):
        if value < self.__DEFAULT_COPIED_SHIPMENT_QTY:
            raise Exception("Copied Quantity Less Default")

        if value > self.shipmentQty:
            raise Exception("Copied Shipment Quantity Over Shipment Quantity")

        self.__copiedShipmentQty = value

    @property
    def isCopyDone(self) -> bool:
        return self.copiedShipmentQty == self.shipmentQty

    def reset_copy(self):
        self.copiedShipmentQty = self.__DEFAULT_COPIED_SHIPMENT_QTY

    @property
    def notCopiedShipmentQty(self) -> int:
        return self.shipmentQty - self.copiedShipmentQty


class OrderFiles:
    # __TBD_DATE = date(2030, 1, 1)
    __TBD_DATE = None
    __FILES_LIMIT = 2

    def __init__(self, files: List[OrderFile]):
        if len(files) == 0:
            raise Exception("No Files")

        self.__files = files

    def append_order_file(self, orderFile: OrderFile):
        if len(self.__files) >= self.__FILES_LIMIT:
            raise Exception(
                f"Number of Files Over Limit: {self.__FILES_LIMIT}"
            )

        self.__files.append(orderFile)

    @property
    def files(self) -> List[OrderFile]:
        return self.__files

    @property
    def orders(self) -> List[Order]:
        orders: List[Order] = list()
        for file in self.__files:
            orders.extend(file.orders)
        return orders

    @property
    def ordersHasNotTBDSPLRow(self) -> List[Order]:
        orders_has_not_tbd_rows: List[Order] = list()
        for file in self.__files:
            orders_has_not_tbd_rows.extend(file.ordersHasNotTBDSPLRow)
        return orders_has_not_tbd_rows

    @property
    def ordersOfKata(self) -> Dict[str, List[Order]]:
        orders_of_kata: Dict[str, List[Order]] = {}
        for file in self.files:
            for kata, orders in file.ordersOfKata.items():
                if kata in orders_of_kata:
                    orders_of_kata[kata].extend(orders)
                else:
                    orders_of_kata[kata] = orders
        return orders_of_kata

    @property
    def releasedQtyOfKata(self) -> Dict[str, int]:
        released_qty_of_kata: Dict[str, int] = {}
        for file in self.files:
            for kata, released_qty in file.releasedQtyOfKata.items():
                if kata in released_qty_of_kata:
                    released_qty_of_kata[kata] += released_qty
                else:
                    released_qty_of_kata[kata] = released_qty
        return released_qty_of_kata

    def apply_shipping_plan(self, pmsFile: PMSFile):
        for kata, pms_rows_of_a_kata in pmsFile.pmsRowsOfKata.items():
            tbd_qty = self.releasedQtyOfKata[kata] - \
                pms_rows_of_a_kata.shipmentQty
            if tbd_qty < 0:
                raise Exception(
                    f"Shipment Quantity Over Released Quantity: {kata}, {tbd_qty}")

            spl_rows: List[SPLRow] = []
            for hin, shipment_qty in pms_rows_of_a_kata.shipmentQtyOfHin.items():
                spl_row = SPLRow(kata=kata,  # type: ignore
                                 hin=hin,
                                 shipmentDate=pms_rows_of_a_kata.shipmentDate,
                                 shipmentQty=shipment_qty,
                                 shipmentWarehouse=pms_rows_of_a_kata.shipmentWarehouse,
                                 isTBD=False)
                spl_rows.append(spl_row)

            if tbd_qty > 0:
                tbd_hin = spl_rows[0].hin
                tbd_spl_row = SPLRow(kata=kata,  # type: ignore
                                     hin=tbd_hin,
                                     shipmentDate=self.__TBD_DATE,
                                     shipmentQty=tbd_qty,
                                     shipmentWarehouse=pms_rows_of_a_kata.shipmentWarehouse,
                                     isTBD=True)
                spl_rows.append(tbd_spl_row)

            self.__spl_rows_to_order(splRows=spl_rows,
                                     orders=self.ordersOfKata[kata])

    @classmethod
    def __spl_rows_to_order(cls,
                            splRows: List[SPLRow],
                            orders: List[Order]):
        for order in orders:
            order.splRows = []

        while True:
            done = cls.__spl_rows_to_order_core(splRows=splRows,
                                                orders=orders)
            if done:
                break

    @classmethod
    def __spl_rows_to_order_core(cls,
                                 splRows: List[SPLRow],
                                 orders: List[Order]) -> bool:
        for order in orders:
            if order.isUpdateDone:
                continue

            for spl_row in splRows:
                if spl_row.isCopyDone:
                    continue

                if order.notUpdatedShipmentQty >= spl_row.notCopiedShipmentQty:
                    # まだ回答を更新していない注残 >= 未引当出荷数
                    # 回答が1注番に全部入る
                    shipment_qty = spl_row.notCopiedShipmentQty
                    spl_row_to_order = SPLRow(kata=spl_row.kata,  # type: ignore
                                              hin=spl_row.hin,
                                              shipmentDate=spl_row.shipmentDate,
                                              shipmentQty=shipment_qty,
                                              shipmentWarehouse=spl_row.shipmentWarehouse,
                                              isTBD=spl_row.isTBD)
                    spl_row.copiedShipmentQty = spl_row.shipmentQty
                    order.updatedReleasedQty += shipment_qty
                    order.splRows.append(spl_row_to_order)
                else:
                    # まだ回答を更新していない注残 < 未引当出荷数
                    # 回答が注番をまたぐ
                    shipment_qty = order.notUpdatedShipmentQty
                    spl_row_to_order = SPLRow(kata=spl_row.kata,  # type: ignore
                                              hin=spl_row.hin,
                                              shipmentDate=spl_row.shipmentDate,
                                              shipmentQty=shipment_qty,
                                              shipmentWarehouse=spl_row.shipmentWarehouse,
                                              isTBD=spl_row.isTBD)
                    spl_row.copiedShipmentQty += shipment_qty
                    order.updatedReleasedQty = order.releasedQty
                    order.splRows.append(spl_row_to_order)

                    return False

        return True
