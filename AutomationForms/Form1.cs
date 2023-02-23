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
            string fileName = string.Format(@"{0}", label1.Text); //���� ���
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
            string workSheetName = "�Ϲ�_";
            try
            {
                if (checkBox1.Checked)
                {
                    workBook = excelApp.Workbooks.Open(label1.Text);
                    workSheetName = "��Ǫ_";
                }
                else
                {
                    workBook = excelApp.Workbooks.Open(label1.Text, Password: passWord);
                }
                // ��ȹ������ �� ����
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
                // ������ ��Ʈ �б�
                itemWorkSheet = workBook.Worksheets.get_Item(2) as Excel.Worksheet;
                if (itemWorkSheet is not null)
                {
                    if (itemWorkSheet.Name.Contains("������") == false)
                    {
                        itemWorkSheet = workBook.Worksheets.get_Item(3) as Excel.Worksheet;
                    }
                }
                string week = comboBox1.Text;
                ReadWriteItem.readItem(itemWorkSheet, out itemData, week, checkBox1.Checked);
                // ���� ��Ʈ �б�
                craftWorkSheet = workBook.Worksheets.get_Item(3) as Excel.Worksheet;
                if (craftWorkSheet is not null) 
                { 
                    if (craftWorkSheet.Name.Contains("����") == false)
                    {
                        craftWorkSheet = workBook.Worksheets.get_Item(4) as Excel.Worksheet;
                    }
                }
                ReadWriteCraft.readCraft(craftWorkSheet, out craftData, week, checkBox1.Checked);
                // ��ȹ�� �ݱ�
                workBook.Close(false);
                // �� ���� ���� ������ �� �Է�
                newWorkBook = excelApp.Workbooks.Add(Type.Missing);
                //workSheet = newWorkBook.Worksheets.Add(Type.Missing) as Worksheet;
                workSheet = newWorkBook.Sheets["Sheet1"] as Excel.Worksheet;
                //workSheet = newWorkBook.ActiveSheet as Excel.Worksheet;
                workSheet.Name = workSheetName + "����";
                writeData(ref workSheet, BMitemData);
                // ������ ��Ʈ ���� �� ������ üũ����Ʈ �ۼ� 
                itemWorkSheet = newWorkBook.Sheets.Add(Type.Missing, After: newWorkBook.Sheets[newWorkBook.Sheets.Count]);
                itemWorkSheet.Name = workSheetName + "������";
                ReadWriteItem.writeItem(ref itemWorkSheet, itemData, checkBox1.Checked);
                // ���� ��Ʈ ���� �� ���� üũ����Ʈ �ۼ�
                craftWorkSheet = newWorkBook.Sheets.Add(Type.Missing, After: newWorkBook.Sheets[newWorkBook.Sheets.Count]);
                craftWorkSheet.Name = workSheetName + "����";
                ReadWriteCraft.writeCraft(ref craftWorkSheet, craftData, checkBox1.Checked);
                // �ڵ�ȭ ��Ʈ, �÷��� ��Ʈ ����
                automateWorkSheet = newWorkBook.Sheets.Add(Type.Missing, After: newWorkBook.Sheets[newWorkBook.Sheets.Count]);
                automateWorkSheet.Name = workSheetName + "�ڵ�ȭ";
                collectionWorkSheet = newWorkBook.Sheets.Add(Type.Missing, After: newWorkBook.Sheets[newWorkBook.Sheets.Count]);
                collectionWorkSheet.Name = workSheetName + "�÷���";
                // �� ���� ����
                String saveFilePath = label6.Text + label4.Text;
                if (checkBox1.Checked)
                {
                    saveFilePath = label6.Text + label4.Text +"_FAFU";
                }
                newWorkBook.SaveAs(saveFilePath);
                // Exit from the application  
                newWorkBook.Close();
                excelApp.Quit();
                MessageBox.Show("BMüũ����Ʈ ���� �Ϸ�.");
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
            // ��ȹ������ �ʿ��� �÷��� ã�Ƽ� �������鼭 �ش� �÷����� ����Ʈ�� ��� return
            int[] colNums;
            // �Ϻ� �븸 ���� ���� ��Ƽ� �Ʒ� list.Add(temp);�� �� ��ŵ
            int JPTWcolNums = 3;
            int weekNameCol = 1;
            // ��Ǫ BM üũ����Ʈ ������ ���� ����־ �ش� �� üũ����Ʈ���� �����ϱ� ���� �߰�
            int ProductName = 5;
            if (checkbox)
            {
                // itemData{(0)Sub Category,(1)��ǰ �̸�,(2)��ȭ,(3)����,(4)���� ����,(5)���� ����,(6)����ǰ,(7)���� ȹ�� ������,(8)�Ǹ� �Ⱓ}
                colNums = new int[] { 5, 6, 7, 8, 11, 12, 14, 15, 17 };
                ProductName = 6;
            }
            else
            {
                colNums = new int[] { 4, 5, 6, 7, 8, 9, 11, 12, 13 };
            }

            // ������ ���� Ȯ��
            Excel.Range rng = workSheet.UsedRange;
            object[,] data = (object[,])rng.Value;
            int dataLength = data.GetLength(0);
            //int rowLength = data.GetLength(1);

            //string category = workSheet.Cells[3, 5].value;
            // ���� Ȯ���Ͽ� ���� �� Ȯ��
            string week = comboBox1.Text;
            string nextWeek = "";
            if (week == "1����")
                nextWeek = "2����";
            else if (week == "2����")
                nextWeek = "3����";
            else if (week == "3����")
                nextWeek = "4����";
            else if (week == "4����")
                nextWeek = "5����";
            else if (week == "5����")
                nextWeek = "6����";
            bool startRow = false;
            for (int i = 1; i <= dataLength; i++)
            {
                // �켱 1������ ������ �´� ���� ã�Ƽ� startRow�� true�� �����.
                // �׷� ������ �´� ����� �����ؼ� �� ������ list�� �־ �ٽ� list<list>�� �ִ´�.
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
                    //���� �Ѿ�� ����Ʈ�� �׸� �ְ� ���� COLUME���� CONTINUE
                    if (data[i, weekNameCol] is string)
                    {
                        temp = data[i, 1] as string;
                        if (temp == nextWeek)
                        {
                            break;
                        }
                    }
                    //��ǰ �̸��� ������ ��ŵ
                    if (data[i, ProductName] is null)
                    {
                        continue;
                    }
                    //JPTW �� ��ŵ
                    if (checkbox) 
                    {
                        if (data[i, JPTWcolNums] is string)
                        {
                            temp = data[i, JPTWcolNums] as string;
                            if (temp == "�Ϻ�" | temp =="�븸")
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
            // ���� B2�� ��¥ ���� ���� ����, �� ���� �� H2���� ���� ����
            workSheet.Cells[2, 2] = $"{label4.Text}";
            Range titleRange = workSheet.Range[workSheet.Cells[2, 2], workSheet.Cells[2, 8]];
            titleRange.Merge();
            basicTitleDesign(titleRange);
            // B4�� ��з�, C4�� �Һз�, D4�� Ȯ�� �׸�, E4�� ���, F4�� �����, G4�� JIRA, H4�� Ȯ�� ����
            Range tempRange;
            workSheet.Cells[4, 2] = $"��з�";
            tempRange = workSheet.Range[workSheet.Cells[4, 2], workSheet.Cells[4, 2]];
            basicTitleDesign(tempRange);

            workSheet.Cells[4, 3] = $"�Һз�";
            tempRange = workSheet.Range[workSheet.Cells[4, 3], workSheet.Cells[4, 3]];
            basicTitleDesign(tempRange);

            workSheet.Cells[4, 4] = $"Ȯ�� �׸�";
            tempRange = workSheet.Range[workSheet.Cells[4, 4], workSheet.Cells[4, 4]];
            basicTitleDesign(tempRange);

            workSheet.Cells[4, 5] = $"���";
            tempRange = workSheet.Range[workSheet.Cells[4, 5], workSheet.Cells[4, 5]];
            basicTitleDesign(tempRange);

            workSheet.Cells[4, 6] = $"�����";
            tempRange = workSheet.Range[workSheet.Cells[4, 6], workSheet.Cells[4, 6]];
            basicTitleDesign(tempRange);

            workSheet.Cells[4, 7] = $"JIRA";
            tempRange = workSheet.Range[workSheet.Cells[4, 7], workSheet.Cells[4, 7]];
            basicTitleDesign(tempRange);

            workSheet.Cells[4, 8] = $"Ȯ�� ����";
            tempRange = workSheet.Range[workSheet.Cells[4, 8], workSheet.Cells[4, 8]];
            basicTitleDesign(tempRange);
            // ������ �ϳ��� �ۼ��ϴ� BasicItemChecklist�� ����, ���� ���� �����ͼ� ���� ���� ���� ������
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
            // itemData{(0)Sub Category,(1)��ǰ �̸�,(2)��ȭ,(3)����,(4)���� ����,(5)���� ����,(6)����ǰ,(7)���� ȹ�� ������,(8)�Ǹ� �Ⱓ}
            // ��з� �ۼ�
            if (itemData[5] == "-")
                workSheet.Cells[writeRow, bigCategory] = $"{itemData[1]}";
            else
                workSheet.Cells[writeRow, bigCategory] = $"{itemData[1]}\n{itemData[5]}";
            // �Һз��� Ȯ�� �׸� �ۼ�
            workSheet.Cells[writeRow, firstCol] = $"��ǰ��";
            workSheet.Cells[writeRow, SecondCol] = $"��ǰ�� ��� Ȯ��";
            writeRow = writeRow + 1;

            workSheet.Cells[writeRow, firstCol] = $"������";
            workSheet.Cells[writeRow, SecondCol] = $"��ǰ ������ ��� Ȯ��";
            writeRow = writeRow + 1;

            workSheet.Cells[writeRow, firstCol] = $"���� ����";
            workSheet.Cells[writeRow, SecondCol] = $"���� ���� Ƚ�� Ȯ�� ({itemData[4]})";
            writeRow = writeRow + 1;

            workSheet.Cells[writeRow, firstCol] = $"����";
            workSheet.Cells[writeRow, SecondCol] = $"��ǰ ���� Ȯ�� ({itemData[2]} {itemData[3]})";
            writeRow = writeRow + 1;
            if (itemData[2] == "���̾�")
            {
                // ���ϸ��� ���� �ݾ� ��� �� ����
                tempString = itemData[3].Replace(",", "");
                int tempInt = Convert.ToInt32(tempString) * 2;
                tempString = Convert.ToString(tempInt);
                workSheet.Cells[writeRow, SecondCol] = $"��{tempString} ���ϸ��� ����";
                // ���� �ʵ�� �� �Ʒ� �ʵ带 ����
                workSheet.Range[workSheet.Cells[writeRow - 1, firstCol], workSheet.Cells[writeRow, firstCol]].Merge();
                // �Ƶ��� ��ǰ�� ���� ���߱� ���� Row�� +1
                writeRow = writeRow + 1;
            }

            workSheet.Cells[writeRow, firstCol] = $"���� �˾� â";
            workSheet.Cells[writeRow, SecondCol] = $"���� �˾�â > ����ǰ Ȯ��";
            writeRow = writeRow + 1;
            workSheet.Cells[writeRow, SecondCol] = $"���� �˾�â > �Ǹ� �Ⱓ Ȯ��\r\n{itemData[8]}";
            writeRow = writeRow + 1;
            tempString = itemData[7].Replace("��� �� �Ʒ� ������ ȹ��", "���� ȹ�� ������");
            workSheet.Cells[writeRow, SecondCol] = $"{tempString}";
            workSheet.Range[workSheet.Cells[writeRow - 2, firstCol], workSheet.Cells[writeRow, firstCol]].Merge();
            writeRow = writeRow + 1;

            workSheet.Cells[writeRow, firstCol] = "������\r\n> ���� ȹ�� ������\r\n��� ��� �� �� ���� Ȯ��";
            tempWords = itemData[7].Split("\n");
            if (itemData[7].Contains("��� �� �Ʒ� ������ ȹ��"))
            {
                tempRow = writeRow;
                foreach (string word in tempWords)
                {
                    if (word.Contains("��� �� �Ʒ� ������ ȹ��")) { continue; }
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
            workSheet.Cells[writeRow, firstCol] = $"���� ��ǰ �� ����";
            workSheet.Cells[writeRow, SecondCol] = $"[���� ��ǰ �� ����] �����Ͽ� �� ���� ��� Ȯ��";
            tempRow = writeRow;
            writeRow = writeRow + 1;
            workSheet.Cells[writeRow, SecondCol] = $"{itemData[7]}";
            writeRow = writeRow + 1;
            workSheet.Cells[writeRow, SecondCol] = "������ ����Ʈ ���� ������ ���������� ��� �� ȹ�� ������ ������/Ŭ����/�ư��ÿ�/��ȥ �ȳ� Ȯ��";
            writeRow = writeRow + 1;
            if (tempString.Contains("��� �� �Ʒ� ������ ȹ��"))
            {
                foreach (string word in tempWords)
                {
                    if (word.Contains("��� �� �Ʒ� ������ ȹ��")) { continue; }
                    workSheet.Cells[writeRow, SecondCol] = word;
                    writeRow = writeRow + 1;
                }
            }
            workSheet.Range[workSheet.Cells[tempRow, firstCol], workSheet.Cells[writeRow - 1, firstCol]].Merge();

            workSheet.Cells[writeRow, firstCol] = $"����";
            workSheet.Cells[writeRow, SecondCol] = $"���� �� ���� Ȯ��";
            writeRow = writeRow + 1;

            workSheet.Cells[writeRow, firstCol] = $"�κ��丮";
            workSheet.Cells[writeRow, SecondCol] = $"��Ű�� ������ ���� �� �κ��丮 �̵� Ȯ��";
            writeRow = writeRow + 1;
            workSheet.Cells[writeRow, SecondCol] = $"��Ű�� ������ �� ����â ������ ��ȹ�� ��ġ�ϴ� ���� Ȯ��";
            writeRow = writeRow + 1;
            workSheet.Cells[writeRow, SecondCol] = $"��Ű�� ������ ���� ��� Ȯ��";
            workSheet.Range[workSheet.Cells[writeRow - 2, firstCol], workSheet.Cells[writeRow, firstCol]].Merge();
            writeRow = writeRow + 1;

            workSheet.Cells[writeRow, firstCol] = $"������ ���";
            workSheet.Cells[writeRow, SecondCol] = $"�κ��丮 ���� ����/����/�̵� �� ��Ű�� ������ ������� �� ������ ���� ���� Ȯ��";
            tempRow = writeRow;
            writeRow = writeRow + 1;
            if (tempString.Contains("��� �� �Ʒ� ������ ȹ��"))
            {
                foreach (string word in tempWords)
                {
                    if (word.Contains("��� �� �Ʒ� ������ ȹ��")) { continue; }
                    workSheet.Cells[writeRow, SecondCol] = word;
                    writeRow = writeRow + 1;
                }
            }
            else
            {
                workSheet.Cells[writeRow, SecondCol] = itemData[1];
                writeRow = writeRow + 1;
            }
            workSheet.Cells[writeRow, SecondCol] = $"+ ������ ��� Ȯ��";
            writeRow = writeRow + 1;
            workSheet.Range[workSheet.Cells[tempRow, firstCol], workSheet.Cells[writeRow - 1, firstCol]].Merge();
            // ��з� �÷� ����
            workSheet.Range[workSheet.Cells[startRow, bigCategory], workSheet.Cells[writeRow - 1, bigCategory]].Merge();
            // ��з� �Һз� Į�� �߾� ����
            //workSheet.Range[workSheet.Cells[startRow, bigCategory], workSheet.Cells[writeRow - 1, firstCol]].HorizontalAlignment = 1;
            // �����¿� �ʺ� fix 
            workSheet.Columns["B:H"].AutoFit();
            workSheet.Rows[$"1:{writeRow}"].AutoFit();
            // ���������� startRow�� writeRow �� �Է�
            startRow = writeRow;
        }
    }
}