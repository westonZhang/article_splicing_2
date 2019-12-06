from file_processing_old_new_mix_2 import File_processing5
from file_split_and_splicing import FileSplitAndSplicing
from utils import Util
import sys


def run_original(path):
    """
    最初的文章拼接程序
    """
    file_splice = FileSplitAndSplicing(path)
    file_splice.run()

def run_mix(path):
    """
    新旧文章的混合拼接程序
    """
    file_splice = File_processing5(path)
    file_splice.run()

if __name__ == '__main__':
    while True:
        util = Util()
        read_path = r'./data/read_path'  # IDE执行
        # read_path = r'../data/read_path'  # 打包执行
        folders = util.get_file_dir(read_path)
        for folder in folders:
            print("-" * 50)
            print(folder)
            print("-" * 50)

            folder_path = read_path + "/" + folder
            if "_a_" in folder:
                run_original(folder_path)
            else:
                run_mix(folder_path)

        while True:
            input_str = input('Press "Y(or y) + Enter" to continue, press "N(or n) + Enter" to exit:')
            if input_str == "Y" or input_str == "y":
                break
            elif input_str == "N" or input_str == "n":
                sys.exit()
            else:
                print("Your input neither Y(y) nor N(n), please input again.")

