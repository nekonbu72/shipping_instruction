from time import sleep
from typing import List, Optional, Tuple

from shipping_instruction.browser import (download_order, shipping_instruction,
                                          upload_spl)
from shipping_instruction.config import (AnsweredOrderFileColumnConfig,
                                         DirConfig, DriverConfig, MRPCConfig,
                                         NewOrderFileColumnConfig)
from shipping_instruction.order import Order, OrderFile, OrderFiles
from shipping_instruction.pdf import merge
from shipping_instruction.pms import PMSFile, PMSFileColumnsConfig
from shipping_instruction.user import User


def read_pms_file() -> PMSFile:
    return PMSFile(path=DirConfig.PMS_FILE_DIR,
                   config=PMSFileColumnsConfig())


def download_answered_order(mrpCConfig: MRPCConfig, user: User) -> str:
    answered_config = DriverConfig(download=DirConfig.ANSWERED_ORDER_DIR)
    answered_file_path = download_order(isNew=False,
                                        driverConfig=answered_config,
                                        mrpCConfig=mrpCConfig,
                                        user=user)
    if answered_file_path is None:
        raise Exception("Answered Order File Download Fail")
    print(f"Answered Order File Path: {answered_file_path}")
    return answered_file_path


def download_new_order(mrpCConfig: MRPCConfig, user: User) -> Optional[str]:
    new_config = DriverConfig(download=DirConfig.NEW_ORDER_DIR)
    new_file_path = download_order(isNew=True,
                                   driverConfig=new_config,
                                   mrpCConfig=mrpCConfig,
                                   user=user)
    if new_file_path is None:
        # 新規受注がゼロの場合もあるためエラーにしない
        print("New Order File Download Fail")
    else:
        print(f"New Order File Path: {new_file_path}")
    return new_file_path


def output_upload_file_wrapper(pmsFile: PMSFile,
                               answeredFilePath: str,
                               newFilePath: Optional[str]) -> Tuple[bool, bool, OrderFiles, Optional[str]]:
    answered_order_file = OrderFile(isNew=False,
                                    path=answeredFilePath,
                                    config=AnsweredOrderFileColumnConfig())

    new_order_file: Optional[OrderFile] = None
    if newFilePath is None:
        print("New Order File Not Found")
    else:
        new_order_file = OrderFile(isNew=True,
                                   path=newFilePath,
                                   config=NewOrderFileColumnConfig())

    order_files = OrderFiles(files=[answered_order_file])

    if not new_order_file is None:
        order_files.append_order_file(orderFile=new_order_file)

    order_files.apply_shipping_plan(pmsFile=pmsFile)

    tyuumon_bangou_prefix = order_files.get_valid_tyuumou_bangou_prefix()
    # if tyuumon_bangou_prefix is None:
    #     raise Exception(" Invalid Tyuumon Bangou")

    answered_done = new_done = False
    for order_file in order_files.files:
        if not order_file.isNew:
            answered_done = order_file.output_upload_file(
                output=DirConfig.ANSWERED_ORDER_OUTPUT_PATH
            )
            if not answered_done:
                print("Answered Order No Update")
            else:
                print(
                    f"Answered Order Output Path: {DirConfig.ANSWERED_ORDER_OUTPUT_PATH}"
                )
        else:
            new_done = order_file.output_upload_file(
                output=DirConfig.NEW_ORDER_OUTPUT_PATH
            )
            if not new_done:
                print("New Order No Update")
            else:
                print(
                    f"New Order Output Path: {DirConfig.NEW_ORDER_OUTPUT_PATH}"
                )

    return (answered_done, new_done, order_files, tyuumon_bangou_prefix)


def upload_spl_wrapper(doAnswered: bool, doNew: bool, user: User):

    answered_done = False
    if doAnswered:
        answered_done = upload_spl(isNew=False,
                                   driverConfig=DriverConfig(download=""),
                                   dirConfig=DirConfig(),
                                   user=user)
    if doAnswered:
        if answered_done:
            print("Answered Order SPL Upload Done")
        else:
            print("Answered Order SPL Upload Fail")

    new_done = False
    if doNew:
        new_done = upload_spl(isNew=True,
                              driverConfig=DriverConfig(download=""),
                              dirConfig=DirConfig(),
                              user=user)
    if doNew:
        if new_done:
            print("New Order SPL Upload Done")
        else:
            print("New Order SPL Upload Fail")

    if (not (doAnswered == answered_done)) \
            or (not (doNew == new_done)):
        raise Exception("Upload Fail")


def shipping_instruction_wrapper(orders: List[Order],
                                 mrpCConfig: MRPCConfig,
                                 user: User):
    shipping_instruction(orders=orders,
                         driverConfig=DriverConfig(download=DirConfig.PDF_DIR),
                         mrpCConfig=mrpCConfig,
                         user=user)


def merge_wrapper(pmsFile: PMSFile):
    if merge(inputDir=DirConfig.PDF_DIR,
             outputBaseDir=DirConfig.PDF_OUTPUT_DIR,
             instructionNumber=pmsFile.instructionNumber) is None:
        raise Exception("PDF Merge Fail")


def main():

    BYE = 5

    print("PMS ファイルを読み込みます")

    pms_file = read_pms_file()

    print(f"このファイルを元に処理を開始します: {pms_file.fileName}")

    mrp_c_config = MRPCConfig(pms_file)
    user = User(jsonPath=DirConfig.USER_JSON_PATH)

    print("注残のファイルをダウンロードします")

    (answered_file_path, new_file_path) = (
        download_answered_order(mrpCConfig=mrp_c_config,
                                user=user),
        download_new_order(mrpCConfig=mrp_c_config,
                           user=user)
    )

    print("納期回答アップロードファイルを作成します")

    (do_answered, do_new, order_files, tyuumon_bangou_prefix) = output_upload_file_wrapper(
        pmsFile=pms_file,
        answeredFilePath=answered_file_path,
        newFilePath=new_file_path
    )

    if tyuumon_bangou_prefix is None:
        print("注文番号の接頭辞が複数混在しているため、処理を中止します")
        print(f"このウィンドウは{BYE}秒後に自動的に閉じます")
        sleep(BYE)
        return

    print(f"注文番号の接頭辞は {tyuumon_bangou_prefix} のみです")

    print("納期回答をアップロードします")

    upload_spl_wrapper(
        doAnswered=do_answered,
        doNew=do_new,
        user=user)

    print("出荷指示を登録します")

    shipping_instruction_wrapper(
        orders=order_files.ordersHasNotTBDSPLRow,
        mrpCConfig=mrp_c_config,
        user=user
    )

    print("出荷指示書の PDF を結合します")

    merge_wrapper(pmsFile=pms_file)

    print(f"処理が完了しました。このウィンドウは{BYE}秒後に自動的に閉じます")
    sleep(BYE)


if __name__ == "__main__":
    main()
