using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace AutomationForms
{
    internal static class ReadWriteCraft
    {
        public static void readCraft(Worksheet workSheet, out object[,] craftData, string week, bool checkbox)
        {
            // 기획서에서 필요한 행부터 데이터가 들어간 range를 통째로 return\
            // F열부터 J열까지가 기획서에서 필요한 열
            int colStart = 1; // A열
            int colEnd = 26; // Z열
            if (checkbox)
            {
                colEnd = 30; // AD열
            }
            Range firstColRange = workSheet.Columns[1];
            object[,] data = (object[,])firstColRange.Value;
            Range usedRange = workSheet.UsedRange;
            object[,] usedData = (object[,])usedRange.Value;
            int dataLength = usedData.GetLength(0);
            int startRow = 1;
            //for (int i = 1; i <= dataLength; i++)
            //{
            //    if (data[i, 1] is null)
            //    {
            //        continue;
            //    }
            //    Type type = data[i, 1].GetType();
            //    if ((type == typeof(string)) && (data[i, 1].ToString().Contains(week)))
            //    {
            //        startRow = i;
            //        break;
            //    }
            //}
            // 주차 확인해서 행 번호를 startRow에 담았기 때문에 dataLength까지 Range를 긁어서 return
            // L, M, N 컬럼이 체크리스트에 안 들어가서 해당 열 제외해야 되지만 방법을 못 찾아 write 완료 후 해당 컬럼 삭제 예정
            Range rng = workSheet.Range[workSheet.Cells[startRow, colStart], workSheet.Cells[dataLength, colEnd]];
            craftData = (object[,])rng.Value;
        }
        public static void writeCraft(ref Worksheet workSheet, object[,] craftData, bool checkbox)
        {
            int writeRow = 1;
            // 주차 텍스트와 카테고리 텍스트 가져와서 write할 때 카테고리 칼럼 날릴 예정
            if (checkbox)
            {
                workSheet.Cells[writeRow, 1] = $"-";
                workSheet.Cells[writeRow, 2] = $"-";
                workSheet.Cells[writeRow, 3] = $"-";
                //
                workSheet.Cells[writeRow, 4] = $"탭";
                workSheet.Cells[writeRow, 5] = $"-"; //테스트 커맨드
                workSheet.Cells[writeRow, 6] = $"제작 Id"; //제작 ID
                workSheet.Cells[writeRow, 7] = $"-"; //제작 아이템 id
                workSheet.Cells[writeRow, 8] = $"-"; //제작 아이템 INT id
                workSheet.Cells[writeRow, 9] = $"제작 아이템";
                workSheet.Cells[writeRow, 10] = $"아이템\n수량";
                workSheet.Cells[writeRow, 11] = $"제작 재료";
                workSheet.Cells[writeRow, 12] = $"수량";
                workSheet.Cells[writeRow, 13] = $"대체 재료";
                workSheet.Cells[writeRow, 14] = $"수량";
                workSheet.Cells[writeRow, 15] = $"제작\n성공\n확률";
                workSheet.Cells[writeRow, 16] = $"대성공\n확률";
                workSheet.Cells[writeRow, 17] = $"대성공 시 획득 아이템";
                workSheet.Cells[writeRow, 18] = $"실패 시 획득 아이템";
                workSheet.Cells[writeRow, 19] = $"실패 시\n획득 수량";
                workSheet.Cells[writeRow, 20] = $"-"; //성공 확률업 ID
                workSheet.Cells[writeRow, 21] = $"성공확률업 적용";
                workSheet.Cells[writeRow, 22] = $"-"; //상자 내 키 아이템 개수
                workSheet.Cells[writeRow, 23] = $"서버 선착순";
                workSheet.Cells[writeRow, 24] = $"월드 선착순";
                workSheet.Cells[writeRow, 25] = $"횟수 제한";
                workSheet.Cells[writeRow, 26] = $"주초기화";
                workSheet.Cells[writeRow, 27] = $"제작 시작일";
                workSheet.Cells[writeRow, 28] = $"제작 종료일";
                workSheet.Cells[writeRow, 29] = $"비고";
                workSheet.Cells[writeRow, 30] = $"아이템삭제일";
            }
            else
            {
                workSheet.Cells[writeRow, 1] = $"-";
                workSheet.Cells[writeRow, 2] = $"-";
                //
                workSheet.Cells[writeRow, 3] = $"탭";
                workSheet.Cells[writeRow, 4] = $"제작 Id";
                workSheet.Cells[writeRow, 5] = $"제작 아이템";
                workSheet.Cells[writeRow, 6] = $"아이템\n수량";
                workSheet.Cells[writeRow, 7] = $"제작 재료";
                workSheet.Cells[writeRow, 8] = $"수량";
                workSheet.Cells[writeRow, 9] = $"대체 재료";
                workSheet.Cells[writeRow, 10] = $"수량";
                workSheet.Cells[writeRow, 11] = $"제작\n성공\n확률";
                // 11,12,13은 지울 컬럼이라서 dummy값 입력
                workSheet.Cells[writeRow, 12] = $"-";
                workSheet.Cells[writeRow, 13] = $"-";
                workSheet.Cells[writeRow, 14] = $"-";
                //
                workSheet.Cells[writeRow, 15] = $"대성공\n확률";
                workSheet.Cells[writeRow, 16] = $"대성공 시 획득 아이템";
                workSheet.Cells[writeRow, 17] = $"실패 시 획득 아이템";
                workSheet.Cells[writeRow, 18] = $"실패 시\n획득 수량";
                workSheet.Cells[writeRow, 19] = $"월드 선착순";
                workSheet.Cells[writeRow, 20] = $"서버 선착순";
                workSheet.Cells[writeRow, 21] = $"횟수 제한";
                workSheet.Cells[writeRow, 22] = $"제작 시작일";
                workSheet.Cells[writeRow, 23] = $"제작 종료일";
                workSheet.Cells[writeRow, 24] = $"성공확률업 적용";
                workSheet.Cells[writeRow, 25] = $"TJ쿠폰\n복구대상";
                workSheet.Cells[writeRow, 26] = $"비고";
            }
            writeRow = writeRow + 1;
            // A열부터 Z열까지 기획서에서 긁어왔기 때문에 동일 크기 range 설정
            int dataLength = craftData.GetLength(0);
            string tempString;
            double? tempDouble;
            bool? tempbool;
            int tempLength = 27;
            if (checkbox)
                tempLength = 31;
            for (int i = 1; i < dataLength + 1; i++)
            {
                for (int j = 1; j < tempLength; j++)
                {
                    if (craftData[i, j] is string)
                    {
                        tempString = craftData[i, j] as string;
                        workSheet.Cells[i + writeRow - 1, j] =  tempString;
                    }
                    else if (craftData[i, j] is double)
                    {
                        tempDouble = craftData[i, j] as double?;
                        workSheet.Cells[i + writeRow - 1, j] = tempDouble;
                    }
                    else if (craftData[i, j] is bool)
                    {
                        tempbool = craftData[i, j] as bool?;
                        workSheet.Cells[i + writeRow - 1, j] = tempbool;
                    }
                    else
                    {
                        workSheet.Cells[i + writeRow - 1, j] = "-";
                    }
                }
            }
            // tablehead에 검은 바탕색에 흰글씨 설정 + 칼럼 ['C']가 "탭"인 row를 뽑아서 검은 바탕, 흰 글씨로 전환
            List <int> tapTextRow = new List<int>();
            Range tapTextColumns = workSheet.Columns["C"];
            if (checkbox)
            {
                tapTextColumns = workSheet.Columns["D"];
            }
            object[,] tapTextData = (object[,])tapTextColumns.Value;
            for (int i = 1; i < dataLength + 1; i++)
            {
                if (tapTextData[i, 1] is string)
                {
                    tempString = tapTextData[i, 1] as string;
                    if (tempString == "탭")
                    {
                        tapTextRow.Add(i);
                    }
                }
            }

            if (checkbox)
            {
                foreach (int row in tapTextRow)
                {
                    workSheet.Cells[row, 31] = $"TJ쿠폰 대상\n 여부(100% 확률이면 대상X)";
                    workSheet.Cells[row, 32] = $"TJ 대체 아이템\n(각인) 설정";
                    workSheet.Cells[row, 33] = $"데이터";
                    workSheet.Cells[row, 34] = $"브로드캐스팅\n 확인(서버,서버군)";
                    workSheet.Cells[row, 35] = $"제작 탭\n 이름 확인";
                    workSheet.Cells[row, 36] = $"제작 리스트에\n 출력 확인";
                    workSheet.Cells[row, 37] = $"제작 재료\n 및 수량 확인";
                    workSheet.Cells[row, 38] = $"대체 재료\n 및 수량 확인";
                    workSheet.Cells[row, 39] = $"제작 재료\n 개수/6=대체 재료 개수";
                    workSheet.Cells[row, 40] = $"성공\n 확률 확인";
                    workSheet.Cells[row, 41] = $"대성공\n 확률 확인";
                    workSheet.Cells[row, 42] = $"성공 후\n 획득 아이템 확인";
                    workSheet.Cells[row, 43] = $"대성공 후\n 획득 아이템 확인";
                    workSheet.Cells[row, 44] = $"실패 후\n 획득 아이템 확인";
                    workSheet.Cells[row, 45] = $"실패 후\n 획득 아이템 개수=제작 재료 개수/10";
                    workSheet.Cells[row, 46] = $"제작 제한\n 종류";
                    workSheet.Cells[row, 47] = $"제작 제한\n 개수";
                    workSheet.Cells[row, 48] = $"제작 \n시작/종료일";
                    workSheet.Cells[row, 49] = $"성공 확률업\n 설정 확인";
                    workSheet.Cells[row, 50] = $"성공 확률업\n 재료 확인";
                    workSheet.Cells[row, 51] = $"성공 확률업\n 재료를 사용하여 제작 확인";
                    workSheet.Cells[row, 52] = $"담당자";
                    workSheet.Cells[row, 53] = $"JIRA";
                    workSheet.Cells[row, 54] = $"비고";
                    Range headRange = workSheet.Range[workSheet.Cells[row, 2], workSheet.Cells[row, 54]];
                    headRange.Font.Color = Color.FromArgb(255, 255, 255);
                    headRange.Interior.Color = Color.FromArgb(0, 0, 0);
                    headRange.Font.Bold = true;
                    headRange.Font.Size = 9;
                    headRange.RowHeight = 54;
                }
                // 좌우 폭 맞춤
                workSheet.Columns["B:G"].AutoFit();
                workSheet.Rows[$"1:{dataLength + 1}"].AutoFit();
                // 체크리스트 22,20,8,7,5,3,2 칼럼 삭제
                workSheet.Columns["V"].Delete(XlDirection.xlToLeft);
                workSheet.Columns["T"].Delete(XlDirection.xlToLeft);
                workSheet.Columns["G:H"].Delete(XlDirection.xlToLeft);
                workSheet.Columns["E"].Delete(XlDirection.xlToLeft);
                workSheet.Columns["B:C"].Delete(XlDirection.xlToLeft);
            }
            else
            {
                foreach (int row in tapTextRow)
                {
                    workSheet.Cells[row, 27] = $"TJ쿠폰 대상\n 여부(100% 확률이면 대상X)";
                    workSheet.Cells[row, 28] = $"TJ 대체 아이템\n(각인) 설정";
                    workSheet.Cells[row, 29] = $"데이터";
                    workSheet.Cells[row, 30] = $"브로드캐스팅\n 확인(서버,서버군)";
                    workSheet.Cells[row, 31] = $"제작 탭\n 이름 확인";
                    workSheet.Cells[row, 32] = $"제작 리스트에\n 출력 확인";
                    workSheet.Cells[row, 33] = $"제작 재료\n 및 수량 확인";
                    workSheet.Cells[row, 34] = $"대체 재료\n 및 수량 확인";
                    workSheet.Cells[row, 35] = $"제작 재료\n 개수/6=대체 재료 개수";
                    workSheet.Cells[row, 36] = $"성공\n 확률 확인";
                    workSheet.Cells[row, 37] = $"대성공\n 확률 확인";
                    workSheet.Cells[row, 38] = $"성공 후\n 획득 아이템 확인";
                    workSheet.Cells[row, 39] = $"대성공 후\n 획득 아이템 확인";
                    workSheet.Cells[row, 40] = $"실패 후\n 획득 아이템 확인";
                    workSheet.Cells[row, 41] = $"실패 후\n 획득 아이템 개수=제작 재료 개수/10";
                    workSheet.Cells[row, 42] = $"제작 제한\n 종류";
                    workSheet.Cells[row, 43] = $"제작 제한\n 개수";
                    workSheet.Cells[row, 44] = $"제작 \n시작/종료일";
                    workSheet.Cells[row, 45] = $"성공 확률업\n 설정 확인";
                    workSheet.Cells[row, 46] = $"성공 확률업\n 재료 확인";
                    workSheet.Cells[row, 47] = $"성공 확률업\n 재료를 사용하여 제작 확인";
                    workSheet.Cells[row, 48] = $"담당자";
                    workSheet.Cells[row, 49] = $"JIRA";
                    workSheet.Cells[row, 50] = $"비고";
                    Range headRange = workSheet.Range[workSheet.Cells[row, 2], workSheet.Cells[row, 50]];
                    headRange.Font.Color = Color.FromArgb(255, 255, 255);
                    headRange.Interior.Color = Color.FromArgb(0, 0, 0);
                    headRange.Font.Bold = true;
                    headRange.Font.Size = 9;
                    headRange.RowHeight = 54;
                }
                // 좌우 폭 맞춤
                workSheet.Columns["B:G"].AutoFit();
                workSheet.Rows[$"1:{dataLength + 1}"].AutoFit();
                // 체크리스트 B 칼럼 삭제
                workSheet.Columns["B"].Delete(XlDirection.xlToLeft);
                // 체크리스트 K,L,M 칼럼 삭제
                workSheet.Columns["K:M"].Delete(XlDirection.xlToLeft);
            }    
        }
    }
}
