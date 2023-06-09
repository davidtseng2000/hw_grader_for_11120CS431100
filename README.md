# hw_grader_for_11120CS431100
A little practice of my "Automatic Grading Program".

# Introduction
When I was a teaching assistant in college, the professor assigned two programming assignments. "When the students completed their programs and submitted them for grading, how can I efficiently grade them in bulk?" To address this issue, I developed two little programs to assist me in grading. One is the **'File Classifier'** and the other is the **'Automatic Grading'**.

# How to use?
## Step1: Folder Setting
After downloading the students' assignments from the learning platform, create an "HW" folder that includes the following items: the students' homeworks, a "data" folder (containing test cases and their solutions), the "File Classifier" program, and the "Automatic Grading Program". To keep track of which students have submitted their assignments and to ensure that no grades are missed during the recording process, also place the grade recording Excel file in the "HW" folder.

在從學習平台下載完學生的作業後，開一個 HW 資料夾，放置**學生的作業**（每個學生的作業都是一個資料夾，裡面放程式和時間複雜度作圖 plot.png）、**data 資料夾**（裡面放 test cases 和 test cases 的解答）、**File Classifier** 和 **Automatic Grading**，為了標記哪些同學有交作業方便後面登記成績沒有漏掉，將登記成績的 **excel** 檔也放到 HW 資料夾中。

↓ HW 資料夾內容示意圖：

![HW 資料夾內容示意圖](https://github.com/davidtseng2000/hw_grader_for_11120CS431100/blob/main/pic/HW_file_all.png)

---

## Step2: **File Classifier.py**
When running "File Classifier.py," it performs the following actions:
1. First, you need to input the keyword for the assignment's file name. For example, if students' assignments for Assignment 4 should include the keyword "hw5_pns", please input "hw5_pns".
2. It organizes the students' programs into three folders based on their file extensions: **.c**, **.cpp**, and **.py.** It also prints out information about files with potential formatting errors, such as incorrect file names (e.g., hw5_abc.py) or files that are not within the specified file types (e.g., hw5_pns.txt).
3. It updates the prepared grade recording Excel file from step 1 and saves it as "_updated.xlsx." If a student has submitted their assignment, the corresponding cell in the student ID column will be highlighted in light blue.

執行 File Classifier.py，它會執行以下幾個動作:
1. 首先你需要輸入作業的檔名的關鍵字，例如在 HW5 中學生的作業應該都會包含關鍵字 "hw5_pns"，這時候請輸入 "hw5_pns"。
2. File Classifier 將學生的程式依照副檔名整理到三個資料夾: .c , .cpp 和 .py，同時會印出格式可能有誤的檔案資訊，例如：檔名錯誤 (hw5_abc.py) 或繳交檔案非上述三類 (hw5_pns.txt)。
3. 更新步驟一中準備的成績登記表 excel 檔並儲存成 "updated.xlsx"，若該同學有繳交作業，則學號部分會塗成淺藍色。

↓ File Classifier 使用畫面：

![File Classifier 使用畫面](https://github.com/davidtseng2000/hw_grader_for_11120CS431100/blob/main/pic/classifier.png)

↓ File Classifier 整理後的資料夾：

![File Classifier 整理後的資料夾](https://github.com/davidtseng2000/hw_grader_for_11120CS431100/blob/main/pic/after_classifier.png)

↓ 原 excel 檔內容：

![原 excel 檔內容](https://github.com/davidtseng2000/hw_grader_for_11120CS431100/blob/main/pic/excel.png)

↓ excel_updated 內容：

![excel_updated 內容](https://github.com/davidtseng2000/hw_grader_for_11120CS431100/blob/main/pic/excel_updated.png)

---

## Step3: **Automatic Grading.py**
To run "Automatic Grading.py," please note that there are some prerequisites:
1. Place "Automatic Grading.py," the "data" folder inside the folder of the assignment you want to grade (.cpp, .c, or .py files).
2. Then execute "Automatic Grading.py" and wait for it to automatically grade the assignments.

執行 Automatic Grading.py ，但注意執行前需要一些前置作業：
1. 將 Automatic Grading.py、data 資料夾丟到要改的作業的資料夾中 (.cpp .c 或 .py)
2. 接著執行 Automatic Grading.py，開始等它自動改完作業！

↓ Automatic Grading 前置作業（範例將 Step3-1 提到的檔案都丟到 ".cpp資料夾" 中）：

![Automatic Grading 前置作業](https://github.com/davidtseng2000/hw_grader_for_11120CS431100/blob/main/pic/before_autograde.png)

↓ Automatic Grading 使用畫面：

![Automatic Grading 使用畫面](https://github.com/davidtseng2000/hw_grader_for_11120CS431100/blob/main/pic/auto_grade.png)
