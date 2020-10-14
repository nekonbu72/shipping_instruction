import subprocess
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
        raise Exception("回答済受注ファイルのダウンロードに失敗しました")

    print(f"回答済受注ファイルをダウンロードしました: {answered_file_path}")
    return answered_file_path


def download_new_order(mrpCConfig: MRPCConfig, user: User) -> Optional[str]:
    new_config = DriverConfig(download=DirConfig.NEW_ORDER_DIR)
    new_file_path = download_order(isNew=True,
                                   driverConfig=new_config,
                                   mrpCConfig=mrpCConfig,
                                   user=user)
    if new_file_path is None:
        # 新規受注がゼロの場合もあるためエラーにしない
        print("新規受注ファイルのダウンロードに失敗しました")
    else:
        print(f"新規受注ファイルをダウンロードしました: {new_file_path}")
    return new_file_path


def output_upload_file_wrapper(pmsFile: PMSFile,
                               answeredFilePath: str,
                               newFilePath: Optional[str]) -> Tuple[bool, bool, OrderFiles, Optional[str]]:
    answered_order_file = OrderFile(isNew=False,
                                    path=answeredFilePath,
                                    config=AnsweredOrderFileColumnConfig())

    new_order_file: Optional[OrderFile] = None
    if newFilePath is None:
        print("新規受注ファイルが見つかりませんでした")
    else:
        new_order_file = OrderFile(isNew=True,
                                   path=newFilePath,
                                   config=NewOrderFileColumnConfig())

    order_files = OrderFiles(files=[answered_order_file])

    if new_order_file is not None:
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
                print("回答済受注に対する納期回答更新はありません")
            else:
                print(
                    f"回答済受注の回答アップロードファイルを作成しました: {DirConfig.ANSWERED_ORDER_OUTPUT_PATH}"
                )
        else:
            new_done = order_file.output_upload_file(
                output=DirConfig.NEW_ORDER_OUTPUT_PATH
            )
            if not new_done:
                print("新規受注に対する納期回答更新はありません")
            else:
                print(
                    f"新規受注の回答アップロードファイルを作成しました: {DirConfig.NEW_ORDER_OUTPUT_PATH}"
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
            print("回答済受注の回答アップロードが完了しました")
        else:
            print("回答済受注の回答アップロードが失敗しました")

    new_done = False
    if doNew:
        new_done = upload_spl(isNew=True,
                              driverConfig=DriverConfig(download=""),
                              dirConfig=DirConfig(),
                              user=user)
    if doNew:
        if new_done:
            print("回答済受注の回答アップロードが完了しました")
        else:
            print("新規受注の回答アップロードが失敗しました")

    if doAnswered != answered_done or doNew != new_done:
        raise Exception("回答アップロードに失敗しました")


def shipping_instruction_wrapper(orders: List[Order],
                                 mrpCConfig: MRPCConfig,
                                 user: User):
    shipping_instruction(orders=orders,
                         driverConfig=DriverConfig(download=DirConfig.PDF_DIR),
                         mrpCConfig=mrpCConfig,
                         user=user)


def merge_wrapper(pmsFile: PMSFile):
    pdf_path = merge(inputDir=DirConfig.PDF_DIR,
                     outputBaseDir=DirConfig.PDF_OUTPUT_DIR,
                     instructionNumber=pmsFile.instructionNumber)

    if pdf_path is None:
        raise Exception("PDF の結合に失敗しました")

    subprocess.Popen(["start", pdf_path], shell=True)


def main():

    BYE = 5

    pms_file = read_pms_file()

    print(f"このファイルをもとに処理を開始します: {pms_file.fileName}")

    mrp_c_config = MRPCConfig(pms_file)
    user = User(jsonPath=DirConfig.USER_JSON_PATH)

    print("")
    print("受注ファイルをダウンロードします")

    (answered_file_path, new_file_path) = (
        download_answered_order(mrpCConfig=mrp_c_config,
                                user=user),
        download_new_order(mrpCConfig=mrp_c_config,
                           user=user)
    )

    print("")
    print("納期回答アップロードファイルを作成します")

    (do_answered, do_new, order_files, tyuumon_bangou_prefix) = output_upload_file_wrapper(
        pmsFile=pms_file,
        answeredFilePath=answered_file_path,
        newFilePath=new_file_path
    )

    if tyuumon_bangou_prefix is None:
        print("")
        print("注文番号の接頭辞が複数混在しているため、処理を中止します")
        # print(f"このウィンドウは{BYE}秒後に自動的に閉じます")
        # sleep(BYE)
        input("エンターキーを押すとこのウィンドウが閉じます")
        return

    print("")
    print(f"注文番号の接頭辞は {tyuumon_bangou_prefix} のみです")

    print("")
    print("納期回答をアップロードします")

    upload_spl_wrapper(
        doAnswered=do_answered,
        doNew=do_new,
        user=user)

    print("")
    print("出荷指示を登録します")

    shipping_instruction_wrapper(
        orders=order_files.ordersHasNotTBDSPLRow,
        mrpCConfig=mrp_c_config,
        user=user
    )

    print("")
    print("出荷指示書の PDF を結合します")

    merge_wrapper(pmsFile=pms_file)

    print("")
    # print(f"処理が完了しました。このウィンドウは{BYE}秒後に自動的に閉じます")
    # sleep(BYE)

    print(f"処理が完了しました")
    input("エンターキーを押すとこのウィンドウが閉じます")


if __name__ == "__main__":
    main()
