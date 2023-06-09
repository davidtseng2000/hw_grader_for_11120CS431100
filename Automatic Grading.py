import unittest
import os
import subprocess
from colorama import init, Fore, Style
import pandas as pd
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows

# 測試資料資料夾路徑
DataPath = "data"
HwPath = "."
Keyword = "pns"

class TestProgram(unittest.TestCase):
    
    def test_program(self):
        
        # 作業關鍵字
        global Keyword

        # 檢查測試資料夾是否存在
        if not os.path.isdir(DataPath):
            raise Exception("找不到測試資料資料夾")
        
        # 取得同學的程式檔案
        program_file = None
        for file_name in os.listdir(HwPath):
            if (Keyword in file_name) and (file_name.endswith(".py") or file_name.endswith(".cpp") or file_name.endswith(".c")):
                program_file = file_name
                student_id =   program_file.split('.')[0].split('_')[2]            
                print(Style.RESET_ALL)
                print(Fore.BLUE + "=" * 16 + "\n" + f"學號: {student_id}" + "\n" + "=" * 16)
                print(Style.RESET_ALL)
                self.score = 0

                for file_name in os.listdir(DataPath):
                    # 取得測試資料路徑
                    if file_name.endswith(".txt"):
                        test_file_path = os.path.join(DataPath, file_name)     
                        # 讀取測試資料
                        with open(test_file_path, 'r') as f:
                            input_data = f.read()
                        
                        try:
                            # 將測試資料代入同學的程式檢查
                            if program_file.endswith(".py"):
                                import_name = os.path.splitext(program_file)[0]
                                program = __import__(import_name)
                                command = ['python', program_file]
                            elif program_file.endswith(".cpp"):
                                # 使用 C++ 編譯並執行程式
                                subprocess.run(['g++', program_file, '-o', 'program'])
                                command = ['./program']
                            else:
                                raise Exception('不支援的程式檔案格式')
                            
                            # 執行程式，將測資代入並取得輸出結果
                            process = subprocess.Popen(
                                command,
                                stdin=subprocess.PIPE,
                                stdout=subprocess.PIPE,
                                stderr=subprocess.PIPE,
                                text=True
                            )
                            output, _ = process.communicate(input=input_data)
                            
                            # 將測資檔名轉換為測試案例編號（例如：1.txt -> Case 1）
                            case_number = os.path.splitext(file_name)[0]
                            case_name = f"Case {case_number}"
                            
                            # 比對輸出結果是否符合預期
                            with self.subTest(case_name=case_name):
                                expected_output_file = os.path.join(DataPath, f"{case_number}.out")
                                # print(expected_output_file)
                                with open(expected_output_file, 'r') as f:
                                    expected_output = f.read().strip()
                                # 將預期輸出與實際輸出進行比對
                                print(f"expected output: {int(expected_output)}, output: {int(output.strip())}")
                                if expected_output in output.strip():
                                    self.assertEqual(expected_output, output.strip())
                                    self.score += 25
                                    print(f"✓ 通過 {case_name}")
                                else:
                                    print(f"✕ 未通過 {case_name}")
                        except:
                            print(Fore.RED + "讀取/編譯檔案時發生錯誤!")
                            print(Style.RESET_ALL)

                if (self.score == 0):
                    print(Fore.RED + f"score: {self.score}")
                else:
                    print(Fore.GREEN + f"score: {self.score}")
                print(Style.RESET_ALL)

if __name__ == '__main__':
    unittest.main()
