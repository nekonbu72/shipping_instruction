import unittest

from shipping_instruction.config import DirConfig, MRPCConfig
from shipping_instruction.main import (
    download_answered_order, download_new_order, main, merge_wrapper,
    output_upload_file_wrapper, read_pms_file, shipping_instruction_wrapper,
    upload_spl_wrapper)
from shipping_instruction.order import OrderFiles
from shipping_instruction.pms import PMSFile
from shipping_instruction.user import User
from shipping_instruction.util import _get_first_file_in_dir


class TestMain(unittest.TestCase):

    def test_read_pms_file(self):
        read_pms_file()

    def test_download_order(self):
        pms_file = read_pms_file()
        mrp_c_config = MRPCConfig(pms_file)
        user = User(DirConfig.USER_JSON_PATH)
        download_answered_order(mrp_c_config, user)
        download_new_order(mrp_c_config, user)

    def test_output_upload_file_wrapper(self) -> OrderFiles:
        pms_file = read_pms_file()

        answered_order_path = _get_first_file_in_dir(
            DirConfig.ANSWERED_ORDER_DIR
        )
        if answered_order_path is None:
            raise Exception("Answered Order File Not Found")

        new_order_path = _get_first_file_in_dir(DirConfig.NEW_ORDER_DIR)

        (_, _, order_files) = output_upload_file_wrapper(
            answered_order_path,
            new_order_path,
            pms_file
        )
        return order_files

    def test_upload_spl_wrapper(self):
        upload_spl_wrapper(False, False, User(DirConfig.USER_JSON_PATH))
        # upload_spl_wrapper(True, False, User(DirConfig.USER_JSON_PATH))
        # upload_spl_wrapper(False, True, User(DirConfig.USER_JSON_PATH))
        # upload_spl_wrapper(False, False, User(DirConfig.USER_JSON_PATH))

    def test_shipping_instruction_wrapper(self):
        pms_file = read_pms_file()
        mrp_c_config = MRPCConfig(pms_file)
        order_files = self.test_output_upload_file_wrapper()
        shipping_instruction_wrapper(
            order_files.ordersHasNotTBDSPLRow,
            mrp_c_config,
            User(DirConfig.USER_JSON_PATH)
        )

    def test_merge_wrapper(self):
        pms_file = read_pms_file()
        merge_wrapper(pms_file)

    def test_main(self):
        main()


if __name__ == "__main__":
    unittest.main()
