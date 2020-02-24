from pathlib import Path
from time import sleep
from typing import List, Optional

# geckodriver, Selenium, Firefox のバージョン対応は下記をチェック
# https://firefox-source-docs.mozilla.org/testing/geckodriver/Support.html
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.webdriver import WebDriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

from shipping_instruction.config import DirConfig, DriverConfig, MRPCConfig
from shipping_instruction.order import Order, OrderFiles
from shipping_instruction.pms import PMSFile
from shipping_instruction.user import User
from shipping_instruction.util import _get_first_file_in_dir

__SPEC = "32268"


class DownloadCompleted():
    __DOWNLOAD_WAIT = 5

    def __init__(self, dir: str):
        sleep(self.__DOWNLOAD_WAIT)
        self.dir = dir

    def __call__(self, driver: WebDriver):
        p = Path(self.dir)
        if not p.is_dir():
            return False

        counter = 0
        for content in p.iterdir():
            counter += 1
            if content.suffix == ".part":
                return False

        return driver if counter >= 1 else False


def download_order(isNew: bool,
                   driverConfig: DriverConfig,
                   mrpCConfig: MRPCConfig,
                   user: User) -> Optional[str]:

    fp = webdriver.FirefoxProfile(profile_directory=driverConfig.profile)
    for key, value in driverConfig.preference.items():
        fp.set_preference(key, value)

    # ファイルをダウンロードする場合、
    # firefox が立ちあがる前に削除しないといけない
    if not driverConfig.delete_handler_files(fp.tempfolder):
        raise Exception("FireFox Profile Temp Folder Error")

    with webdriver.Firefox(
        firefox_profile=fp,
        firefox_binary=driverConfig.firefox,
        executable_path=driverConfig.geckodriver,
        service_log_path=driverConfig.log
    ) as driver:

        wait = WebDriverWait(driver, 10)

        driver.get(user.URL)

        wait.until(
            EC.presence_of_element_located((By.NAME, "sei_login"))
        ).send_keys(user.SSO_ID)

        wait.until(
            EC.presence_of_element_located((By.NAME, "sei_passwd"))
        ).send_keys(user.SSO_PASSWORD)

        wait.until(
            EC.presence_of_element_located((By.NAME, "login"))
        ).submit()

        wait.until(
            EC.frame_to_be_available_and_switch_to_it("fr_menu")
        )

        wait.until(
            EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "body > table:nth-child(6) > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(1) > nobr:nth-child(1) > a:nth-child(1)")
            )
        ).click()

        wait.until(
            EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "body > table:nth-child(6) > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(1) > table:nth-child(2) > tbody:nth-child(1) > tr:nth-child(6) > td:nth-child(1) > a:nth-child(1)")
            )
        ).click()

        driver.switch_to.parent_frame()
        wait.until(
            EC.frame_to_be_available_and_switch_to_it("fr_main")
        )

        # 「検索」をクリック
        if not isNew:
            wait.until(
                EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, "#menu2 > a:nth-child(1)")
                )
            ).click()
            wait.until(
                EC.presence_of_element_located(
                    (By.XPATH, "//*[text()=\" 検索対象\"]")  # 「検」の前の半角スペースに注意
                )
            )

        wait.until(
            EC.presence_of_element_located((By.NAME, "xmrp_bu_c_rfc_2"))
        ).clear()

        sleep(1)

        wait.until(
            EC.presence_of_element_located((By.NAME, "xmrp_bu_c_rfc_2"))
        ).send_keys(mrpCConfig.MRPC)

        wait.until(
            EC.presence_of_element_located(
                (By.NAME, "keiyaku_kaisya_cd_rfc3"))
        ).send_keys(__SPEC)

        wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "#xsotype_k_chk_4-1"))
        ).click()

        wait.until(
            EC.element_to_be_clickable((By.NAME, "btn_submit"))
        ).click()

        wait.until(
            EC.element_to_be_clickable((By.LINK_TEXT, "ダウンロード(XLS)"))
        ).click()

        download_dir = driverConfig.download
        if download_dir is None:
            return None

        WebDriverWait(driver, 30).until(
            DownloadCompleted(download_dir)
        )

        return _get_first_file_in_dir(download_dir)


def upload_spl(isNew: bool,
               driverConfig: DriverConfig,
               dirConfig: DirConfig,
               user: User) -> bool:

    fp = webdriver.FirefoxProfile(profile_directory=driverConfig.profile)
    for key, value in driverConfig.preference.items():
        fp.set_preference(key, value)

    with webdriver.Firefox(
        firefox_profile=fp,
        firefox_binary=driverConfig.firefox,
        executable_path=driverConfig.geckodriver,
        service_log_path=driverConfig.log,
    ) as driver:

        wait = WebDriverWait(driver, 10)

        driver.get(user.URL)

        wait.until(
            EC.presence_of_element_located((By.NAME, "sei_login"))
        ).send_keys(user.SSO_ID)

        wait.until(
            EC.presence_of_element_located((By.NAME, "sei_passwd"))
        ).send_keys(user.SSO_PASSWORD)

        wait.until(
            EC.presence_of_element_located((By.NAME, "login"))
        ).submit()

        wait.until(
            EC.frame_to_be_available_and_switch_to_it("fr_menu")
        )

        wait.until(
            EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "body > table:nth-child(6) > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(1) > nobr:nth-child(1) > a:nth-child(1)")
            )
        ).click()

        wait.until(
            EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "body > table:nth-child(6) > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(1) > table:nth-child(2) > tbody:nth-child(1) > tr:nth-child(7) > td:nth-child(1) > a:nth-child(1)")
            )
        ).click()

        driver.switch_to.parent_frame()
        wait.until(
            EC.frame_to_be_available_and_switch_to_it("fr_main")
        )

        # 「アップロード(更新)」をクリック
        if not isNew:
            wait.until(
                EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, "#menu2 > a:nth-child(1)")
                )
            ).click()
            wait.until(
                EC.presence_of_element_located(
                    (By.XPATH, "//*[text()=\"本機能では、登録済の納期回答を一括変更します。\"]")
                )
            )

        UPLOAD_FILE_PATH = dirConfig.NEW_ORDER_OUTPUT_PATH if isNew else dirConfig.ANSWERED_ORDER_OUTPUT_PATH
        upload_file_path = str(Path(UPLOAD_FILE_PATH).resolve())

        wait.until(
            EC.presence_of_element_located((By.NAME, "pms_upfile"))
        ).send_keys(upload_file_path)

        # test
        wait.until(
            EC.presence_of_element_located((By.NAME, "btn_submit"))
        ).click()

        try:
            wait.until(
                EC.presence_of_element_located(
                    # 「以」の前の改行に注意
                    (By.XPATH, "//*[text()=\"\n以下のデータを登録しました。\"]")
                )
            )
            return True
        except(TimeoutException):
            return False


def shipping_instruction(orders: List[Order],
                         driverConfig: DriverConfig,
                         mrpCConfig: MRPCConfig,
                         user: User):

    fp = webdriver.FirefoxProfile(profile_directory=driverConfig.profile)
    for key, value in driverConfig.preference.items():
        fp.set_preference(key, value)

    # ファイルをダウンロードする場合、
    # firefox が立ちあがる前に削除しないといけない
    if not driverConfig.delete_handler_files(fp.tempfolder):
        raise Exception("FireFox Profile Temp Folder Error")

    with webdriver.Firefox(
        firefox_profile=fp,
        firefox_binary=driverConfig.firefox,
        executable_path=driverConfig.geckodriver,
        service_log_path=driverConfig.log,
    ) as driver:

        wait = WebDriverWait(driver, 10)

        driver.get(user.URL)

        wait.until(
            EC.presence_of_element_located((By.NAME, "sei_login"))
        ).send_keys(user.SSO_ID)

        wait.until(
            EC.presence_of_element_located((By.NAME, "sei_passwd"))
        ).send_keys(user.SSO_PASSWORD)

        wait.until(
            EC.presence_of_element_located((By.NAME, "login"))
        ).submit()

        for order in orders:
            for spl_row in order.notTBDSPLRows:

                wait.until(
                    EC.frame_to_be_available_and_switch_to_it("fr_menu")
                )

                wait.until(
                    EC.element_to_be_clickable(
                        (By.CSS_SELECTOR, "body > table:nth-child(6) > tbody:nth-child(1) > tr:nth-child(2) > td:nth-child(1) > nobr:nth-child(1) > a:nth-child(1)")
                    )
                ).click()

                wait.until(
                    EC.element_to_be_clickable(
                        (By.CSS_SELECTOR, "body > table:nth-child(6) > tbody:nth-child(1) > tr:nth-child(2) > td:nth-child(1) > table:nth-child(2) > tbody:nth-child(1) > tr:nth-child(2) > td:nth-child(1) > a:nth-child(1)")
                    )
                ).click()

                driver.switch_to.parent_frame()
                wait.until(
                    EC.frame_to_be_available_and_switch_to_it("fr_main")
                )

                wait.until(
                    EC.presence_of_element_located(
                        (By.NAME, "xmrp_bu_c_rf_01"))
                ).clear()

                sleep(1)

                wait.until(
                    EC.presence_of_element_located(
                        (By.NAME, "xmrp_bu_c_rf_01"))
                ).send_keys(mrpCConfig.MRPC)

                wait.until(
                    EC.presence_of_element_located((By.NAME, "kaito_noki"))
                ).send_keys(str(spl_row.shipmentDate))

                wait.until(
                    EC.presence_of_element_located(
                        (By.NAME, "pms_to_kaito_noki")
                    )
                ).send_keys(str(spl_row.shipmentDate))

                wait.until(
                    EC.presence_of_element_located((By.NAME, "moku_noki_nn"))
                ).send_keys(str(spl_row.shipmentDate))

                wait.until(
                    EC.presence_of_element_located(
                        (By.NAME, "pms_to_moku_noki_nn")
                    )
                ).send_keys(str(spl_row.shipmentDate))

                wait.until(
                    EC.presence_of_element_located((By.NAME, "seiban2"))
                ).send_keys(order.orderNumber)

                wait.until(
                    EC.presence_of_element_located((By.NAME, "xitm_no_rfc_01"))
                ).send_keys(spl_row.hin)

                wait.until(
                    EC.element_to_be_clickable((By.NAME, "btn_submit"))
                ).submit()

                # 出荷指示の新規登録画面で回答納期が検索できないと、ここで詰まる
                try:
                    wait.until(
                        EC.presence_of_element_located(
                            (By.NAME, "load_cd_rfc_2_0"))
                    ).send_keys(mrpCConfig.TSUMI_BASYO)
                except (TimeoutException):
                    raise Exception("SPL Not Found")

                wait.until(
                    EC.element_to_be_clickable((By.NAME, "updchk_0"))
                ).click()

                sleep(3)

                wait.until(
                    EC.element_to_be_clickable((By.NAME, "btn_submit"))
                ).submit()

                wait.until(
                    EC.presence_of_element_located(
                        # 「以」の前の改行に注意
                        (By.XPATH, "//*[text()=\"\n以下のデータを登録しますか？\"]")
                    )
                )

                wait.until(
                    EC.element_to_be_clickable((By.NAME, "btn_submit"))
                ).submit()

                wait.until(
                    EC.presence_of_element_located(
                        # 「以」の前の改行に注意
                        (By.XPATH, "//*[text()=\"\nデータを登録しました。\"]")
                    )
                )

                # driver.switch_to.parent_frame()
                driver.get(user.URL)
