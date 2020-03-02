from time import sleep
from typing import List, Optional, Tuple

from shipping_instruction.browser import (
    download_order, shipping_instruction, upload_spl)
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


def output_upload_file_wrapper(pms_file: PMSFile,
                               answered_file_path: str,
                               new_file_path: Optional[str]) -> Tuple[bool, bool, OrderFiles]:
    answered_order_file = OrderFile(isNew=False,
                                    path=answered_file_path,
                                    config=AnsweredOrderFileColumnConfig())

    new_order_file: Optional[OrderFile] = None
    if new_file_path is None:
        print("New Order File Not Found")
    else:
        new_order_file = OrderFile(isNew=True,
                                   path=new_file_path,
                                   config=NewOrderFileColumnConfig())

    order_files = OrderFiles([answered_order_file])

    if not new_order_file is None:
        order_files.append_order_file(new_order_file)

    order_files.apply_shipping_plan(pms_file)

    answered_done = new_done = False
    for order_file in order_files.files:
        if not order_file.isNew:
            answered_done = order_file.output_upload_file(
                DirConfig.ANSWERED_ORDER_OUTPUT_PATH
            )
            if not answered_done:
                print("Answered Order No Update")
            else:
                print(
                    f"Answered Order Output Path: {DirConfig.ANSWERED_ORDER_OUTPUT_PATH}"
                )
        else:
            new_done = order_file.output_upload_file(
                DirConfig.NEW_ORDER_OUTPUT_PATH
            )
            if not new_done:
                print("New Order No Update")
            else:
                print(
                    f"New Order Output Path: {DirConfig.NEW_ORDER_OUTPUT_PATH}"
                )

    return (answered_done, new_done, order_files)


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
    if merge(DirConfig.PDF_DIR,
             DirConfig.PDF_OUTPUT_DIR,
             pmsFile.instructionNumber) is None:
        raise Exception("PDF Merge Fail")


def main():

    print("PMS ファイルを読み込みます")

    pms_file = read_pms_file()

    print(f"このファイルを元に処理を開始します: {pms_file.fileName}")

    mrp_c_config = MRPCConfig(pms_file)
    user = User(DirConfig.USER_JSON_PATH)

    print("注残のファイルをダウンロードします")

    (answered_file_path, new_file_path) = (
        download_answered_order(mrp_c_config,
                                user),
        download_new_order(mrp_c_config,
                           user)
    )

    print("納期回答アップロードファイルを作成します")

    (do_answered, do_new, order_files) = output_upload_file_wrapper(
        pms_file,
        answered_file_path,
        new_file_path
    )

    print("納期回答をアップロードします")

    upload_spl_wrapper(
        do_answered,
        do_new,
        user)

    print("出荷指示を登録します")

    shipping_instruction_wrapper(
        order_files.ordersHasNotTBDSPLRow,
        mrp_c_config,
        user
    )

    print("出荷指示書の PDF を結合します")

    merge_wrapper(pms_file)

    BYE = 5
    print(f"処理が完了しました。このウィンドウは{BYE}秒後に自動的に閉じます")
    sleep(BYE)


if __name__ == "__main__":
    main()
