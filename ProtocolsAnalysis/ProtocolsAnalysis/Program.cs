using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ProtocolsAnalysis
{
    class Program
    {
        class Protocol
        {
            public string UID { get; set; }
            public bool IsRight { get; set; }
        }

        static void Main(string[] args)
        {
            var list = new List<Protocol>();
            var index = 1;
            var isRight = false;

            Console.WriteLine("Введите путь к файлу с протоколами ответов студентов.");
            var pathToFile = Console.ReadLine();

            Excel.Application exlApp = new Excel.Application();
            Excel.Workbook exlWb = exlApp.Workbooks.Open(pathToFile);

            var countSheets = exlWb.Sheets.Count;

            for (int i = 1; i <= countSheets; i++)
            {
                Console.WriteLine("Обработка протокола №{0}", i);

                Excel._Worksheet exlWs = exlWb.Sheets[i];
                Excel.Range exlRange = exlWs.UsedRange;

                for (int j = 1; j < exlRange.Rows.Count; j++)
                {
                    if (exlRange.Cells[j, 1].Value != null)
                    {
                        if (exlRange.Cells[j, 1].Value.ToString().Trim() == "Вопрос:")
                        {
                            if (index < 10)
                            {
                                isRight = (exlRange.Cells[j, 2].Value.ToString().Substring(3, (exlRange.Cells[j, 2].Value.ToString().Trim().Length - 4)) == "Правильный ответ") ? true : false;
                            }
                            else
                            {
                                isRight = (exlRange.Cells[j, 2].Value.ToString().Substring(4, (exlRange.Cells[j, 2].Value.ToString().Trim().Length - 5)) == "Правильный ответ") ? true : false;
                            }

                            list.Add(new Protocol()
                            {
                                UID = exlRange.Cells[j + 1, 2].Value.ToString().Trim(),
                                IsRight = isRight
                            });
                            index++;
                        }
                    }
                }

                Marshal.ReleaseComObject(exlRange);
                Marshal.ReleaseComObject(exlWs);
            }

            exlWb.Close();
            exlApp.Quit();

            Marshal.ReleaseComObject(exlWb);
            Marshal.ReleaseComObject(exlApp);

            var uniqueUID = list.Select(x => x.UID).Distinct().ToArray();

            Excel.Application excel = new Excel.Application();
            Excel.Workbook workbook = excel.Workbooks.Add();

            Excel._Worksheet worksheet = workbook.Sheets[1];
            Excel.Range range = worksheet.UsedRange;

            range.Cells[1, 1].Value = "УИК";
            range.Cells[1, 2].Value = "Правильный ответ";
            range.Cells[1, 2].Style.WrapText = true;
            range.Cells[1, 3].Value = "Неправильный ответ";
            range.Cells[1, 3].Style.WrapText = true;

            worksheet.Columns[1].ColumnWidth = 20;
            worksheet.Columns[2].ColumnWidth = 25;
            worksheet.Columns[3].ColumnWidth = 25;

            int row = 2;

            int success = 0;
            int fail = 0;

            bool isSuccess = false;
            bool isFail = false;

            foreach (var uid in uniqueUID)
            {
                var search = list.FindAll(x => x.UID == uid);

                range.Cells[row, 1].Value = uid;

                foreach (var item in search)
                {
                    range.Cells[row, (item.IsRight) ? 2 : 3].Value = "x";

                    switch (item.IsRight)
                    {
                        case true:
                            if (!isSuccess)
                            {
                                success++;
                                isSuccess = true;
                            }
                            break;
                        case false:
                            if (!isFail)
                            {
                                fail++;
                                isFail = true;
                            }
                            break;
                    }
                }

                isSuccess = false;
                isFail = false;

                row++;
            }

            range.Cells[row, 1].Value = "ИТОГО";
            range.Cells[row, 2].Value = success;
            range.Cells[row, 3].Value = fail;

            Console.WriteLine("Введите путь для сохранения файла с результатами обработки.");
            var pathToSaveFile = Console.ReadLine();

            workbook.SaveAs(string.Concat(pathToSaveFile, "NewBook.xlsx"));

            workbook.Close();
            excel.Quit();

            Marshal.ReleaseComObject(range);
            Marshal.ReleaseComObject(worksheet);

            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(excel);

            GC.Collect();
            GC.WaitForPendingFinalizers();

            Console.WriteLine("Обработка протоколов завершена.");
            Console.WriteLine("Для выхода из приложения нажмите любую клавишу.");
            Console.ReadKey();
        }
    }
}
