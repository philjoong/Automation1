using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Reflection.Metadata;
using System.Text;
using System.Threading.Tasks;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace AutomationForms
{
    internal static class ReadWriteItem
    {
        public static void readItem(Worksheet workSheet, out object[,] itemData, string week, bool checkbox)
        {
            // 기획서에서 필요한 행부터 데이터가 들어간 range를 통째로 return\
            // F열부터 J열까지가 기획서에서 필요한 열
            int colStart = 5; // E열
            int colEnd = 9; // I열
            if (checkbox)
            {
                colStart = 5; // E열
                colEnd = 15; // O열
            }
            Range firstColRange = workSheet.Columns[1];
            object[,] data = (object[,])firstColRange.Value;
            Range usedRange = workSheet.UsedRange;
            object[,] usedData = (object[,])usedRange.Value;
            int dataLength = usedData.GetLength(0);
            int startRow = 1;
            for (int i = 1; i <= dataLength; i++)
            {
                if (data[i, 1] is null)
                {
                    continue;
                }
                Type type = data[i, 1].GetType();
                if ((type == typeof(string)) && (data[i, 1].ToString().Contains(week)))
                {
                    startRow = i;
                    break;
                }
            }
            // 주차 확인해서 행 번호를 startRow에 담았기 때문에 dataLength까지 Range를 긁어서 return
            Range rng = workSheet.Range[workSheet.Cells[startRow, colStart], workSheet.Cells[dataLength, colEnd]];
            itemData = (object[,])rng.Value;
        }
        public static void writeItem(ref Worksheet workSheet, object[,] itemData, bool checkbox)
        {
            int writeRow = 1;
            if (checkbox)
            {
                workSheet.Cells[writeRow, 2] = $"int ID";
                workSheet.Cells[writeRow, 3] = $"Item Name";
                workSheet.Cells[writeRow, 4] = $"";
                workSheet.Cells[writeRow, 5] = $"아이템명";
                workSheet.Cells[writeRow, 6] = $"효과";
                workSheet.Cells[writeRow, 7] = $"아이템\n삭제 예정일";
                workSheet.Cells[writeRow, 8] = $"아이콘명";
                workSheet.Cells[writeRow, 9] = $"창고 보관 설정";
                workSheet.Cells[writeRow, 10] = $"거래소 등록";
                workSheet.Cells[writeRow, 11] = $"판매 가능";
                workSheet.Cells[writeRow, 12] = $"삭제 가능";
                workSheet.Cells[writeRow, 13] = $"아이템명";
                workSheet.Cells[writeRow, 14] = $"상세 설명";
                workSheet.Cells[writeRow, 15] = $"아이템 삭제일";
                workSheet.Cells[writeRow, 16] = $"구성품_삭제\n일>= 패키지";
                workSheet.Cells[writeRow, 17] = $"아이콘";
                workSheet.Cells[writeRow, 18] = $"등급";
                workSheet.Cells[writeRow, 19] = $"무게=0";
                workSheet.Cells[writeRow, 20] = $"창고 불가";
                workSheet.Cells[writeRow, 21] = $"삭제 불가";
                workSheet.Cells[writeRow, 22] = $"거래 불가";
                workSheet.Cells[writeRow, 23] = $"판매 불가\n(매입 상인)";
                workSheet.Cells[writeRow, 24] = $"결과_\n아이템사용";
                workSheet.Cells[writeRow, 25] = $"사용 효과\n(구성품/ 효과)";
                workSheet.Cells[writeRow, 26] = $"아이템일괄사용";
                workSheet.Cells[writeRow, 27] = $"퀵슬롯\n자동 사용";
                workSheet.Cells[writeRow, 28] = $"소환권\n사용 효과 확인";
                workSheet.Cells[writeRow, 29] = $"소환권일 경우\n가챠 테이블\n확인 필요";
                workSheet.Cells[writeRow, 30] = $"데이터";
                workSheet.Cells[writeRow, 31] = $"지라";
                workSheet.Cells[writeRow, 32] = $"비고";
            }
            else
            {
                workSheet.Cells[writeRow, 2] = $"int ID";
                workSheet.Cells[writeRow, 3] = $"Item Name";
                workSheet.Cells[writeRow, 4] = $"아이템명";
                workSheet.Cells[writeRow, 5] = $"효과";
                workSheet.Cells[writeRow, 6] = $"아이템\n삭제 예정일";
                workSheet.Cells[writeRow, 7] = $"아이템명";
                workSheet.Cells[writeRow, 8] = $"상세 설명";
                workSheet.Cells[writeRow, 9] = $"아이템 삭제일";
                workSheet.Cells[writeRow, 10] = $"구성품_삭제\n일>= 패키지";
                workSheet.Cells[writeRow, 11] = $"아이콘";
                workSheet.Cells[writeRow, 12] = $"등급";
                workSheet.Cells[writeRow, 13] = $"무게=0";
                workSheet.Cells[writeRow, 14] = $"창고 불가";
                workSheet.Cells[writeRow, 15] = $"삭제 불가";
                workSheet.Cells[writeRow, 16] = $"거래 불가";
                workSheet.Cells[writeRow, 17] = $"판매 불가\n(매입 상인)";
                workSheet.Cells[writeRow, 18] = $"결과_\n아이템사용";
                workSheet.Cells[writeRow, 19] = $"사용 효과\n(구성품/ 효과)";
                workSheet.Cells[writeRow, 20] = $"아이템일괄사용";
                workSheet.Cells[writeRow, 21] = $"퀵슬롯\n자동 사용";
                workSheet.Cells[writeRow, 22] = $"소환권\n사용 효과 확인";
                workSheet.Cells[writeRow, 23] = $"소환권일 경우\n가챠 테이블\n확인 필요";
                workSheet.Cells[writeRow, 24] = $"데이터";
                workSheet.Cells[writeRow, 25] = $"지라";
                workSheet.Cells[writeRow, 26] = $"비고";
            }

            int lastColNum;
            if (checkbox)
            {
                lastColNum = 32;
            }
            else
            {
                lastColNum = 26;
            }
            // tablehead에 검은 바탕색에 흰글씨 설정
            Range headRange = workSheet.Range[workSheet.Cells[writeRow, 2], workSheet.Cells[writeRow, lastColNum]];
            headRange.Font.Color = Color.FromArgb(255, 255, 255);
            headRange.Interior.Color = Color.FromArgb(0, 0, 0);
            headRange.Font.Bold = true;
            headRange.Font.Size = 9;
            writeRow = writeRow + 1;

            // range 붙여넣기
            int dataLength = itemData.GetLength(0);
            // F열부터 J열까지 기획서에서 긁어왔기 때문에 동일 크기 range 설정
            string tempString;
            double? tempDouble;
            bool? tempBool;
            int tempLength = 6;
            if (checkbox)
                tempLength = 12;
            for (int i = 1; i < dataLength + 1; i++)
            {
                for (int j = 1; j < tempLength; j++)
                {
                    if (itemData[i, j] is string)
                    {
                        tempString = itemData[i, j] as string;
                        workSheet.Cells[i + writeRow - 1, j + 1] = tempString;
                    }
                    else if (itemData[i, j] is double)
                    {
                        tempDouble = itemData[i, j] as double?;
                        workSheet.Cells[i + writeRow - 1, j + 1] = tempDouble;
                    }
                    else if (itemData[i, j] is bool)
                    {
                        tempBool = itemData[i, j] as bool?;
                        workSheet.Cells[i + writeRow - 1, j + 1] = tempBool;
                    }
                    else
                    {
                        workSheet.Cells[i + writeRow - 1, j + 1] = "-";
                    }
                }
            }
            // 좌우 폭 맞춤
            workSheet.Columns["B:F"].AutoFit();
            workSheet.Rows[$"1:{dataLength + 1}"].AutoFit();
        }
    }
}
