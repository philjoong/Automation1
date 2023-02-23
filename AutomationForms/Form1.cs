using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System;
using Microsoft.Office.Interop.Excel;
using System.Security.AccessControl;
using System.Diagnostics;

namespace AutomationForms
{
    using static System.Net.Mime.MediaTypeNames;
    using static System.Net.WebRequestMethods;
    using Range = Excel.Range;
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
            label4.Text = DateTime.Now.ToString("yyyyMMdd") + "_BM";
            label6.Text = AppDomain.CurrentDomain.BaseDirectory;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog OFD = new OpenFileDialog();
            if (OFD.ShowDialog() == DialogResult.OK)
            {
                label1.Text = OFD.FileName;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string fileName = string.Format(@"{0}", label1.Text); //파일 경로
            makeChecklist(fileName);
        }

        private void makeChecklist(string fileName)
        {
            object passwordValue = System.Reflection.Missing.Value;
            string passWord = textBox1.Text;

            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = false;
            excelApp.DisplayAlerts = false;
            Workbook? workBook = null;
            Workbook? newWorkBook = null;
            Worksheet? workSheet = null;
            Worksheet? itemWorkSheet = null;
            Worksheet? craftWorkSheet = null;
            Worksheet? automateWorkSheet = null;
            Worksheet? collectionWorkSheet = null;
            object[,] itemData;
            object[,] craftData;
            string workSheetName = "일반_";
            try
            {
                if (checkBox1.Checked)
                {
                    workBook = excelApp.Workbooks.Open(label1.Text);
                    workSheetName = "파푸_";
                }
                else
                {
                    workBook = excelApp.Workbooks.Open(label1.Text, Password: passWord);
                }
                // 기획서에서 값 추출
                workSheet = workBook.Worksheets.get_Item(1) as Excel.Worksheet;
                if (workSheet is not null)
                {
                    if (workSheet.Name.Contains("BM") == false)
                    {
                        workSheet = workBook.Worksheets.get_Item(2) as Excel.Worksheet;
                    }
                }
                List<List<String>> BMitemData;
                readData(workSheet, out BMitemData, checkBox1.Checked);
                // 아이템 시트 읽기
                itemWorkSheet = workBook.Worksheets.get_Item(2) as Excel.Worksheet;
                if (itemWorkSheet is not null)
                {
                    if (itemWorkSheet.Name.Contains("아이템") == false)
                    {
                        itemWorkSheet = workBook.Worksheets.get_Item(3) as Excel.Worksheet;
                    }
                }
                string week = comboBox1.Text;
                ReadWriteItem.readItem(itemWorkSheet, out itemData, week, checkBox1.Checked);
                // 제작 시트 읽기
                craftWorkSheet = workBook.Worksheets.get_Item(3) as Excel.Worksheet;
                if (craftWorkSheet is not null) 
                { 
                    if (craftWorkSheet.Name.Contains("제작") == false)
                    {
                        craftWorkSheet = workBook.Worksheets.get_Item(4) as Excel.Worksheet;
                    }
                }
                ReadWriteCraft.readCraft(craftWorkSheet, out craftData, week, checkBox1.Checked);
                // 기획서 닫기
                workBook.Close(false);
                // 새 파일 만들어서 추출한 값 입력
                newWorkBook = excelApp.Workbooks.Add(Type.Missing);
                //workSheet = newWorkBook.Worksheets.Add(Type.Missing) as Worksheet;
                workSheet = newWorkBook.Sheets["Sheet1"] as Excel.Worksheet;
                //workSheet = newWorkBook.ActiveSheet as Excel.Worksheet;
                workSheet.Name = workSheetName + "상점";
                writeData(ref workSheet, BMitemData);
                // 아이템 시트 생성 및 아이템 체크리스트 작성 
                itemWorkSheet = newWorkBook.Sheets.Add(Type.Missing, After: newWorkBook.Sheets[newWorkBook.Sheets.Count]);
                itemWorkSheet.Name = workSheetName + "아이템";
                ReadWriteItem.writeItem(ref itemWorkSheet, itemData, checkBox1.Checked);
                // 제작 시트 생성 및 제작 체크리스트 작성
                craftWorkSheet = newWorkBook.Sheets.Add(Type.Missing, After: newWorkBook.Sheets[newWorkBook.Sheets.Count]);
                craftWorkSheet.Name = workSheetName + "제작";
                ReadWriteCraft.writeCraft(ref craftWorkSheet, craftData, checkBox1.Checked);
                // 자동화 시트, 컬렉션 시트 생성
                automateWorkSheet = newWorkBook.Sheets.Add(Type.Missing, After: newWorkBook.Sheets[newWorkBook.Sheets.Count]);
                automateWorkSheet.Name = workSheetName + "자동화";
                collectionWorkSheet = newWorkBook.Sheets.Add(Type.Missing, After: newWorkBook.Sheets[newWorkBook.Sheets.Count]);
                collectionWorkSheet.Name = workSheetName + "컬렉션";
                // 새 파일 저장
                String saveFilePath = label6.Text + label4.Text;
                if (checkBox1.Checked)
                {
                    saveFilePath = label6.Text + label4.Text +"_FAFU";
                }
                newWorkBook.SaveAs(saveFilePath);
                // Exit from the application  
                newWorkBook.Close();
                excelApp.Quit();
                MessageBox.Show("BM체크리스트 생성 완료.");
            }
            catch (Exception e)
            {                
                MessageBox.Show($"{e}");
            }
            finally
            {
                ReleaseExcelObject(workSheet);
                ReleaseExcelObject(workBook);
                ReleaseExcelObject(newWorkBook);
                ReleaseExcelObject(itemWorkSheet);
                ReleaseExcelObject(craftWorkSheet);
                ReleaseExcelObject(automateWorkSheet);
                ReleaseExcelObject(collectionWorkSheet);
                ReleaseExcelObject(excelApp);
            }
        }

        private void readData(Worksheet workSheet, out List<List<String>> BMitemData, bool checkbox)
        {
            BMitemData = new List<List<String>>();
            // 기획서에서 필요한 컬럼을 찾아서 루프돌면서 해당 컬럼값을 리스트에 담아 return
            int[] colNums;
            // 일본 대만 행을 따로 담아서 아래 list.Add(temp);할 때 스킵
            int JPTWcolNums = 3;
            int weekNameCol = 1;
            // 파푸 BM 체크리스트 마지막 행이 비어있어서 해당 행 체크리스트에서 제거하기 위해 추가
            int ProductName = 5;
            if (checkbox)
            {
                // itemData{(0)Sub Category,(1)상품 이름,(2)재화,(3)가격,(4)구매 제한,(5)스텝 정보,(6)지급품,(7)개별 획득 아이템,(8)판매 기간}
                colNums = new int[] { 5, 6, 7, 8, 11, 12, 14, 15, 17 };
                ProductName = 6;
            }
            else
            {
                colNums = new int[] { 4, 5, 6, 7, 8, 9, 11, 12, 13 };
            }

            // 데이터 끝줄 확인
            Excel.Range rng = workSheet.UsedRange;
            object[,] data = (object[,])rng.Value;
            int dataLength = data.GetLength(0);
            //int rowLength = data.GetLength(1);

            //string category = workSheet.Cells[3, 5].value;
            // 주차 확인하여 시작 행 확인
            string week = comboBox1.Text;
            string nextWeek = "";
            if (week == "1주차")
                nextWeek = "2주차";
            else if (week == "2주차")
                nextWeek = "3주차";
            else if (week == "3주차")
                nextWeek = "4주차";
            else if (week == "4주차")
                nextWeek = "5주차";
            else if (week == "5주차")
                nextWeek = "6주차";
            bool startRow = false;
            for (int i = 1; i <= dataLength; i++)
            {
                // 우선 1열에서 주차에 맞는 행을 찾아서 startRow를 true로 만든다.
                // 그럼 주차에 맞는 행부터 시작해서 행 단위로 list에 넣어서 다시 list<list>에 넣는다.
                if (startRow != true)
                {
                    if (data[i, 1] is null)
                    {
                        continue;
                    }
                    Type type = data[i, 1].GetType();
                    if ((type == typeof(string)) && (data[i, 1].ToString().Contains(week)))
                    {
                        startRow = true;
                    }
                }
                if (startRow == true)
                {
                    List<String> list = new List<String>();
                    string? temp;
                    //주차 넘어가면 리스트에 그만 넣고 다음 COLUME으로 CONTINUE
                    if (data[i, weekNameCol] is string)
                    {
                        temp = data[i, 1] as string;
                        if (temp == nextWeek)
                        {
                            break;
                        }
                    }
                    //상품 이름이 없으면 스킵
                    if (data[i, ProductName] is null)
                    {
                        continue;
                    }
                    //JPTW 행 스킵
                    if (checkbox) 
                    {
                        if (data[i, JPTWcolNums] is string)
                        {
                            temp = data[i, JPTWcolNums] as string;
                            if (temp == "일본" | temp =="대만")
                            {
                                continue;
                            }
                        }
                    }


                    foreach(int colNum in colNums)
                    {    
                        if (data[i, colNum] is string)
                        {
                            temp = data[i, colNum] as string;             
                            list.Add(temp);
                        }
                        else if(data[i, colNum] is double)
                        {
                            temp = Convert.ToString(data[i, colNum]);
                            list.Add(temp);
                        }
                        else
                        {
                            list.Add("-");
                        }
                    }
                    BMitemData.Add(list);
                }
            }
        }

        private void writeData(ref Worksheet workSheet, List<List<String>> itemData)
        {
            // 먼저 B2에 날짜 적고 검은 바탕, 흰 글자 및 H2까지 병합 설정
            workSheet.Cells[2, 2] = $"{label4.Text}";
            Range titleRange = workSheet.Range[workSheet.Cells[2, 2], workSheet.Cells[2, 8]];
            titleRange.Merge();
            basicTitleDesign(titleRange);
            // B4에 대분류, C4에 소분류, D4에 확인 항목, E4에 결과, F4에 담당자, G4에 JIRA, H4에 확인 내용
            Range tempRange;
            workSheet.Cells[4, 2] = $"대분류";
            tempRange = workSheet.Range[workSheet.Cells[4, 2], workSheet.Cells[4, 2]];
            basicTitleDesign(tempRange);

            workSheet.Cells[4, 3] = $"소분류";
            tempRange = workSheet.Range[workSheet.Cells[4, 3], workSheet.Cells[4, 3]];
            basicTitleDesign(tempRange);

            workSheet.Cells[4, 4] = $"확인 항목";
            tempRange = workSheet.Range[workSheet.Cells[4, 4], workSheet.Cells[4, 4]];
            basicTitleDesign(tempRange);

            workSheet.Cells[4, 5] = $"결과";
            tempRange = workSheet.Range[workSheet.Cells[4, 5], workSheet.Cells[4, 5]];
            basicTitleDesign(tempRange);

            workSheet.Cells[4, 6] = $"담당자";
            tempRange = workSheet.Range[workSheet.Cells[4, 6], workSheet.Cells[4, 6]];
            basicTitleDesign(tempRange);

            workSheet.Cells[4, 7] = $"JIRA";
            tempRange = workSheet.Range[workSheet.Cells[4, 7], workSheet.Cells[4, 7]];
            basicTitleDesign(tempRange);

            workSheet.Cells[4, 8] = $"확인 내용";
            tempRange = workSheet.Range[workSheet.Cells[4, 8], workSheet.Cells[4, 8]];
            basicTitleDesign(tempRange);
            // 아이템 하나씩 작성하는 BasicItemChecklist를 만듦, 시작 행을 가져와서 다음 시작 행을 리턴함
            int startRow = 5;
            foreach(List<String> _itemData in itemData)
            {
                if (_itemData.Count > 0)
                {
                    basicItemChecklist(ref workSheet, ref startRow, _itemData);
                }
            }
        }

        protected void textBox1_TextChanged(object sender, EventArgs e)
        {
          
        }

        private void basicTitleDesign(Range range)
        {
            range.Font.Color = Color.FromArgb(255, 255, 255);
            range.Interior.Color = Color.FromArgb(0, 0, 0);
            range.Font.Bold = true;
        }

        private static void ReleaseExcelObject(object? obj)
        {
            try
            {
                if (obj != null)
                {
                    Marshal.ReleaseComObject(obj);
                    obj = null;
                }
            }
            catch (Exception ex)
            {
                obj = null;
                throw ex;
            }
            finally
            {
                GC.Collect();
            }
        }

        public void basicItemChecklist(ref Worksheet workSheet, ref int startRow, List<String> itemData)
        {
            int writeRow = startRow;
            int tempRow;
            string[] tempWords;
            string tempString;
            int bigCategory = 2;
            int firstCol = 3;
            int SecondCol = 4;
            // itemData{(0)Sub Category,(1)상품 이름,(2)재화,(3)가격,(4)구매 제한,(5)스텝 정보,(6)지급품,(7)개별 획득 아이템,(8)판매 기간}
            // 대분류 작성
            if (itemData[5] == "-")
                workSheet.Cells[writeRow, bigCategory] = $"{itemData[1]}";
            else
                workSheet.Cells[writeRow, bigCategory] = $"{itemData[1]}\n{itemData[5]}";
            // 소분류와 확인 항목 작성
            workSheet.Cells[writeRow, firstCol] = $"상품명";
            workSheet.Cells[writeRow, SecondCol] = $"상품명 출력 확인";
            writeRow = writeRow + 1;

            workSheet.Cells[writeRow, firstCol] = $"아이콘";
            workSheet.Cells[writeRow, SecondCol] = $"상품 아이콘 출력 확인";
            writeRow = writeRow + 1;

            workSheet.Cells[writeRow, firstCol] = $"구매 제한";
            workSheet.Cells[writeRow, SecondCol] = $"구매 제한 횟수 확인 ({itemData[4]})";
            writeRow = writeRow + 1;

            workSheet.Cells[writeRow, firstCol] = $"가격";
            workSheet.Cells[writeRow, SecondCol] = $"상품 가격 확인 ({itemData[2]} {itemData[3]})";
            writeRow = writeRow + 1;
            if (itemData[2] == "다이아")
            {
                // 마일리지 지급 금액 계산 후 삽입
                tempString = itemData[3].Replace(",", "");
                int tempInt = Convert.ToInt32(tempString) * 2;
                tempString = Convert.ToString(tempInt);
                workSheet.Cells[writeRow, SecondCol] = $"ㄴ{tempString} 마일리지 지급";
                // 가격 필드와 그 아래 필드를 병합
                workSheet.Range[workSheet.Cells[writeRow - 1, firstCol], workSheet.Cells[writeRow, firstCol]].Merge();
                // 아데나 상품과 행을 맞추기 위해 Row에 +1
                writeRow = writeRow + 1;
            }

            workSheet.Cells[writeRow, firstCol] = $"구매 팝업 창";
            workSheet.Cells[writeRow, SecondCol] = $"구매 팝업창 > 구성품 확인";
            writeRow = writeRow + 1;
            workSheet.Cells[writeRow, SecondCol] = $"구매 팝업창 > 판매 기간 확인\r\n{itemData[8]}";
            writeRow = writeRow + 1;
            tempString = itemData[7].Replace("사용 시 아래 아이템 획득", "개별 획득 아이템");
            workSheet.Cells[writeRow, SecondCol] = $"{tempString}";
            workSheet.Range[workSheet.Cells[writeRow - 2, firstCol], workSheet.Cells[writeRow, firstCol]].Merge();
            writeRow = writeRow + 1;

            workSheet.Cells[writeRow, firstCol] = "돋보기\r\n> 개별 획득 아이템\r\n목록 출력 및 상세 정보 확인";
            tempWords = itemData[7].Split("\n");
            if (itemData[7].Contains("사용 시 아래 아이템 획득"))
            {
                tempRow = writeRow;
                foreach (string word in tempWords)
                {
                    if (word.Contains("사용 시 아래 아이템 획득")) { continue; }
                    workSheet.Cells[writeRow, SecondCol] = word;
                    writeRow = writeRow + 1;
                }
                workSheet.Range[workSheet.Cells[tempRow, firstCol], workSheet.Cells[writeRow - 1, firstCol]].Merge();
            }
            else
            {
                workSheet.Cells[writeRow, SecondCol] = itemData[7];
                writeRow = writeRow + 1;
            }
            workSheet.Cells[writeRow, firstCol] = $"지급 상품 상세 정보";
            workSheet.Cells[writeRow, SecondCol] = $"[지급 상품 상세 정보] 선택하여 상세 정보 출력 확인";
            tempRow = writeRow;
            writeRow = writeRow + 1;
            workSheet.Cells[writeRow, SecondCol] = $"{itemData[7]}";
            writeRow = writeRow + 1;
            workSheet.Cells[writeRow, SecondCol] = "아이템 리스트 우측 돋보기 아이콘으로 사용 시 획득 가능한 아이템/클래스/아가시온/집혼 안내 확인";
            writeRow = writeRow + 1;
            if (tempString.Contains("사용 시 아래 아이템 획득"))
            {
                foreach (string word in tempWords)
                {
                    if (word.Contains("사용 시 아래 아이템 획득")) { continue; }
                    workSheet.Cells[writeRow, SecondCol] = word;
                    writeRow = writeRow + 1;
                }
            }
            workSheet.Range[workSheet.Cells[tempRow, firstCol], workSheet.Cells[writeRow - 1, firstCol]].Merge();

            workSheet.Cells[writeRow, firstCol] = $"구매";
            workSheet.Cells[writeRow, SecondCol] = $"구매 후 지급 확인";
            writeRow = writeRow + 1;

            workSheet.Cells[writeRow, firstCol] = $"인벤토리";
            workSheet.Cells[writeRow, SecondCol] = $"패키지 아이템 구매 시 인벤토리 이동 확인";
            writeRow = writeRow + 1;
            workSheet.Cells[writeRow, SecondCol] = $"패키지 아이템 상세 정보창 설명이 기획과 일치하는 것을 확인";
            writeRow = writeRow + 1;
            workSheet.Cells[writeRow, SecondCol] = $"패키지 아이템 정상 출력 확인";
            workSheet.Range[workSheet.Cells[writeRow - 2, firstCol], workSheet.Cells[writeRow, firstCol]].Merge();
            writeRow = writeRow + 1;

            workSheet.Cells[writeRow, firstCol] = $"아이템 사용";
            workSheet.Cells[writeRow, SecondCol] = $"인벤토리 슬롯 정렬/삭제/이동 후 패키지 아이템 사용했을 때 아이템 정상 습득 확인";
            tempRow = writeRow;
            writeRow = writeRow + 1;
            if (tempString.Contains("사용 시 아래 아이템 획득"))
            {
                foreach (string word in tempWords)
                {
                    if (word.Contains("사용 시 아래 아이템 획득")) { continue; }
                    workSheet.Cells[writeRow, SecondCol] = word;
                    writeRow = writeRow + 1;
                }
            }
            else
            {
                workSheet.Cells[writeRow, SecondCol] = itemData[1];
                writeRow = writeRow + 1;
            }
            workSheet.Cells[writeRow, SecondCol] = $"+ 아이템 사용 확인";
            writeRow = writeRow + 1;
            workSheet.Range[workSheet.Cells[tempRow, firstCol], workSheet.Cells[writeRow - 1, firstCol]].Merge();
            // 대분류 컬럼 통합
            workSheet.Range[workSheet.Cells[startRow, bigCategory], workSheet.Cells[writeRow - 1, bigCategory]].Merge();
            // 대분류 소분류 칼럼 중앙 정렬
            //workSheet.Range[workSheet.Cells[startRow, bigCategory], workSheet.Cells[writeRow - 1, firstCol]].HorizontalAlignment = 1;
            // 상하좌우 너비 fix 
            workSheet.Columns["B:H"].AutoFit();
            workSheet.Rows[$"1:{writeRow}"].AutoFit();
            // 마지막으로 startRow에 writeRow 값 입력
            startRow = writeRow;
        }
    }
}