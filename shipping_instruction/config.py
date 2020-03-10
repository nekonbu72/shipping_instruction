from getpass import getuser
from pathlib import Path
from typing import Any, Dict, Optional

# 循環参照を防ぐ
# from pms import PMSFile
from shipping_instruction.util import _init_dir


class DriverConfig:
    __MIME_TYPES = ["application/epub+zip",
                    "application/gzip",
                    "application/java-archive",
                    "application/json",
                    "application/ld+json",
                    "application/msword",
                    "application/octet-stream",
                    "application/ogg",
                    "application/pdf",
                    "application/rtf",
                    "application/vnd.amazon.ebook",
                    "application/vnd.apple.installer+xml",
                    "application/vnd.mozilla.xul+xml",
                    "application/vnd.ms-excel",
                    "application/vnd.ms-fontobject",
                    "application/vnd.ms-powerpoint",
                    "application/vnd.oasis.opendocument.presentation",
                    "application/vnd.oasis.opendocument.spreadsheet",
                    "application/vnd.oasis.opendocument.text",
                    "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    "application/vnd.visio",
                    "application/x-7z-compressed",
                    "application/x-abiword",
                    "application/x-bzip",
                    "application/x-bzip2",
                    "application/x-csh",
                    "application/x-freearc",
                    "application/xhtml+xml",
                    "application/xml",
                    "application/x-rar-compressed",
                    "application/x-sh",
                    "application/x-shockwave-flash",
                    "application/x-tar",
                    "application/zip",
                    "appliction/php",
                    "audio/aac",
                    "audio/midi audio/x-midi",
                    "audio/mpeg",
                    "audio/ogg",
                    "audio/wav",
                    "audio/webm",
                    "font/otf",
                    "font/ttf",
                    "font/woff",
                    "font/woff2",
                    "image/bmp",
                    "image/gif",
                    "image/jpeg",
                    "image/png",
                    "image/svg+xml",
                    "image/tiff",
                    "image/vnd.microsoft.icon",
                    "image/webp",
                    "text/calendar",
                    "text/css",
                    "text/csv",
                    "text/html",
                    "text/javascript",
                    "text/javascript",
                    "text/plain",
                    "text/xml",
                    "video/3gpp",
                    "video/3gpp2",
                    "video/mp2t",
                    "video/mpeg",
                    "video/ogg",
                    "video/webm",
                    "video/x-msvideo"]

    __USER_DEFINED = 2

    __HANDLERS = ["mimeTypes.rdf", "handlers.json"]

    def __init__(self,
                 profile: Optional[str] = None,
                 firefox: str = "C:\\Program Files\\Mozilla Firefox\\firefox.exe",
                 geckodriver: str = "geckodriver.exe",
                 log: str = "log",
                 download: str = ""):
        if profile is None:
            self.profile = self.__get_profile_dir()

        self.firefox = firefox

        self.geckodriver = geckodriver

        self.log = self.__setup_log_file_path(log)

        self.preference: Dict[str, Any] = {}

        if download == "":
            self.download = download
            return

        self.download = self.__setup_download_dir(download)

        self.preference = {"browser.download.useDownloadDir": True,
                           "browser.helperApps.neverAsk.saveToDisk": ",".join(self.__MIME_TYPES),
                           "browser.download.folderList": self.__USER_DEFINED,
                           "browser.download.lastDir": "",
                           "browser.download.dir": self.download}

    def delete_handler_files(self, tempfolder: str) -> bool:
        tempfolder_p = Path(tempfolder)
        if not tempfolder_p.is_dir():
            return False

        cnt = 0
        for inner in tempfolder_p.iterdir():
            if cnt >= 1:
                return False

            for name in self.__HANDLERS:
                tgt = inner.joinpath(name)
                if tgt.is_file():
                    tgt.unlink()

            cnt += 1

        return True

    @staticmethod
    def __get_profile_dir() -> Optional[str]:
        user = getuser()
        PROFILE_DIR = f"C:\\Users\\{user}\\AppData\\Roaming\\Mozilla\\Firefox\\Profiles\\"
        p = Path(PROFILE_DIR)
        if p.is_dir():
            for content in p.iterdir():
                if content.is_dir():
                    if "default" in content.name.lower():
                        return str(content)
        return None

    @staticmethod
    def __setup_download_dir(dir: str) -> str:
        # この実装だとダウンロード先のフォルダは毎回リフレッシュされるので注意
        download_dir = _init_dir(dir, True)
        # download は selenium ライブラリ側でエラーが出せないため
        if download_dir is None:
            raise Exception("Download Dir Not Exist")

        return str(Path(download_dir).resolve())

    @staticmethod
    def __setup_log_file_path(dir: str) -> str:
        log_dir = _init_dir(dir, False)
        if log_dir is None:
            return ""

        return str(Path(log_dir).joinpath("geckodriver.log").resolve())


class OrderFileColumnConfingBase:
    SHEET: Optional[str] = None

    TYUUSYAKU_VAL: Optional[str] = None

    RELEASED_VAL: Optional[str] = None
    TEISEKI_RIYUU_VAL: Optional[str] = None
    DELETE_VAL: Optional[str] = None

    SAKUJO_F: Optional[int] = None
    JUTYUU_ID: Optional[int] = None
    JUTYUU_KOUSHIN_NICHIJI: Optional[int] = None
    JUTYUU_ORDER_BANGOU: Optional[int] = None
    JUTYUU_KEY: Optional[int] = None
    TYUUMON_BANGOU: Optional[int] = None
    MEISAI_BANGOU: Optional[int] = None
    KOKYAKU_TYUUMON_BANGOU: Optional[int] = None
    JUTYUU_SYURUI_KUBUN: Optional[int] = None
    YOTAKU_F: Optional[int] = None
    KATABAN: Optional[int] = None
    JUTYUU_SUU: Optional[int] = None
    TEHAI_BI: Optional[int] = None
    KIBOU_NOUKI: Optional[int] = None
    KEIYAKUSAKI_MEI: Optional[int] = None
    JUYOUSAKI_MEI: Optional[int] = None
    NIJI_JUYOUSAKI_MEI: Optional[int] = None
    OKURISAKI_JUUSYO: Optional[int] = None
    BIKOU: Optional[int] = None
    JUTYUU_RECORD_KOUSHIN_BI: Optional[int] = None
    NOUKI_KAITOU_HDR_RECORD_KOUSHIN_BI: Optional[int] = None
    NOUKI_KAITOU_DTL_RECORD_KOUSHIN_BI: Optional[int] = None
    NOUKI_KAITOU_DID: Optional[int] = None
    OIBAN: Optional[int] = None
    SYUKKA_STATUS: Optional[int] = None
    HINBAN: Optional[int] = None
    KAITOU_SUU: Optional[int] = None
    KAITOU_SYUKKA_BI: Optional[int] = None
    MOKUHYOU_NOUKI: Optional[int] = None
    HIKARI_MRP_KAITOU_NOUKI: Optional[int] = None
    SPEC_TYOKUSOU: Optional[int] = None
    NAMAMUGI: Optional[int] = None
    SYUKKA_SOUKO: Optional[int] = None
    YOTAKUSAKI_SOUKO: Optional[int] = None
    TEISEI_RIYUU_C: Optional[int] = None
    TEISEI_RIYUU_SYOSAI_NAIYOU: Optional[int] = None
    TYUUSYAKU: Optional[int] = None
    G_SHIZAI_TYUUMON_BANGOU: Optional[int] = None
    SPEC_JUTYUU_BANGOU: Optional[int] = None
    RENBAN: Optional[int] = None
    GENSANCHI: Optional[int] = None


class NewOrderFileColumnConfig(OrderFileColumnConfingBase):
    SHEET = "Sheet1"

    # TYUUSYAKU_VAL = "部材支給SPEC 免税品(製造部材)"
    TYUUSYAKU_VAL = ""

    JUTYUU_ID = 0
    JUTYUU_ORDER_BANGOU = 1
    JUTYUU_KEY = 2
    TYUUMON_BANGOU = 3
    MEISAI_BANGOU = 4
    KOKYAKU_TYUUMON_BANGOU = 5
    JUTYUU_SYURUI_KUBUN = 6
    YOTAKU_F = 7
    KATABAN = 8
    JUTYUU_SUU = 9
    TEHAI_BI = 10
    KIBOU_NOUKI = 11
    KEIYAKUSAKI_MEI = 12
    JUYOUSAKI_MEI = 13
    NIJI_JUYOUSAKI_MEI = 14
    OKURISAKI_JUUSYO = 15
    BIKOU = 16
    HINBAN = 17
    KAITOU_SUU = 18
    KAITOU_SYUKKA_BI = 19
    MOKUHYOU_NOUKI = 20
    HIKARI_MRP_KAITOU_NOUKI = 21
    SPEC_TYOKUSOU = 22
    NAMAMUGI = 23
    SYUKKA_SOUKO = 24
    YOTAKUSAKI_SOUKO = 25
    TYUUSYAKU = 26


class AnsweredOrderFileColumnConfig(OrderFileColumnConfingBase):
    SHEET = "Sheet1"

    # TYUUSYAKU_VAL = "部材支給SPEC 免税品(製造部材)"
    TYUUSYAKU_VAL = ""

    RELEASED_VAL = "Released"  # Answered だけ
    TEISEKI_RIYUU_VAL = "更新"    # Answered だけ
    DELETE_VAL = "削除"    # Answered だけ

    SAKUJO_F = 0
    JUTYUU_ID = 1
    JUTYUU_KOUSHIN_NICHIJI = 2
    JUTYUU_ORDER_BANGOU = 3
    JUTYUU_KEY = 4
    TYUUMON_BANGOU = 5
    MEISAI_BANGOU = 6
    KOKYAKU_TYUUMON_BANGOU = 7
    JUTYUU_SYURUI_KUBUN = 8
    YOTAKU_F = 9
    KATABAN = 10
    JUTYUU_SUU = 11
    TEHAI_BI = 12
    KIBOU_NOUKI = 13
    KEIYAKUSAKI_MEI = 14
    JUYOUSAKI_MEI = 15
    NIJI_JUYOUSAKI_MEI = 16
    OKURISAKI_JUUSYO = 17
    BIKOU = 18
    JUTYUU_RECORD_KOUSHIN_BI = 19
    NOUKI_KAITOU_HDR_RECORD_KOUSHIN_BI = 20
    NOUKI_KAITOU_DTL_RECORD_KOUSHIN_BI = 21
    NOUKI_KAITOU_DID = 22
    OIBAN = 23
    SYUKKA_STATUS = 24
    HINBAN = 25
    KAITOU_SUU = 26
    KAITOU_SYUKKA_BI = 27
    MOKUHYOU_NOUKI = 28
    HIKARI_MRP_KAITOU_NOUKI = 29
    SPEC_TYOKUSOU = 30
    NAMAMUGI = 31
    SYUKKA_SOUKO = 32
    YOTAKUSAKI_SOUKO = 33
    TEISEI_RIYUU_C = 34
    TEISEI_RIYUU_SYOSAI_NAIYOU = 35
    TYUUSYAKU = 36
    G_SHIZAI_TYUUMON_BANGOU = 37
    SPEC_JUTYUU_BANGOU = 38
    RENBAN = 39
    GENSANCHI = 40


class PMSFileColumnsConfig:
    SHIPMENT_DATE_FORMAT_VAL = "%Y/%m/%d"

    INSTRUCTION_NUMBER = 0
    SHIPMENT_WAREHOUSE = 1
    SHIPMENT_DATE = 3
    KATA = 8
    HIN = 9
    SHIPMENT_QTY = 15


class DirConfig:
    USER_JSON_PATH = "user.json"

    PMS_FILE_DIR = "input"

    ANSWERED_ORDER_DIR = "download\\answered"
    NEW_ORDER_DIR = "download\\new"

    ANSWERED_ORDER_OUTPUT_PATH = "output\\answered.xls"
    NEW_ORDER_OUTPUT_PATH = "output\\new.xls"

    PDF_DIR = "download\\pdf"
    PDF_OUTPUT_DIR = "output\\pdf"

    ERROR_SCREENSHOT_DIR = "error"


class MRPCConfig:
    def __init__(self, pms_file):
        if pms_file.headCharOfShipmentWarehouse == "N":
            self.MRPC = "40"
            self.TSUMI_BASYO = "N05"
        elif pms_file.headCharOfShipmentWarehouse == "E":
            self.MRPC = "20"
            self.TSUMI_BASYO = "E09"
        else:
            raise Exception(
                f"Unknown Shipment Warehouse: {pms_file.headCharOfShipmentWarehouse}"
            )
