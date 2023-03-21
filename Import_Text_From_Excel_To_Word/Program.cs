using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeOpenXml;

namespace Import_Text_From_Excel_To_Word
{
    internal class Program
    {
        /*
        static void Main()
        {
            string excelPath = @"D:\UTA\data.xlsx";
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Инициализация ExcelPackage
            using (ExcelPackage package = new ExcelPackage(new FileInfo(excelPath)))
            {   
                // Открытие первого листа
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                // Определение размеров листа
                int rows = worksheet.Dimension.Rows;
                int columns = worksheet.Dimension.Columns;

                // Чтение данных из ячеек и вывод в виде двумерного массива
                for (int row = 1; row <= rows; row++)
                {
                    for (int col = 1; col <= columns; col++)
                    {
                        object value = worksheet.Cells[row, col].Value;
                        Console.Write(value == null ? "null" : value.ToString());
                        Console.Write("\t");
                    }
                    Console.WriteLine();
                    Console.ReadLine();
                }
            }
        }
        */

        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            string templatePath = @"D:\UTA\temp.docx";
            string dataPath = @"D:\UTA\data.xlsx";
            string outputPath = @"D:\UTA\output";

            ProcessDocuments(templatePath, dataPath, outputPath);
        }

        static void ProcessDocuments(string templatePath, string dataPath, string outputPath)
        {
            using (var package = new ExcelPackage(new FileInfo(dataPath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                Random random = new Random();

                DateTime startDate = new DateTime(2015, 1, 1);
                DateTime endDate = new DateTime(2022, 2, 1);

                int rowCount = worksheet.Dimension.End.Row;
                for (int i = 1; i <= rowCount; i++)
                {
                    string decNumber = worksheet.Cells[i, 1].Text;
                    string docNumber = worksheet.Cells[i, 2].Text;

                    DateTime docDate = RandomDate(startDate, endDate, random);

                    string outputFileName = Path.Combine(outputPath, $"{decNumber}.docx");
                    ReplaceWordsInTemplate(templatePath, outputFileName, i, decNumber, docNumber, docDate);
                }
            }
        }

        static DateTime RandomDate(DateTime startDate, DateTime endDate, Random random)
        {
            int range = (endDate - startDate).Days;
            int randomDay = random.Next(range);
            return startDate.AddDays(randomDay);
        }

        static void ReplaceWordsInTemplate(string templatePath, string outputFileName, int count, string decNumber, string docNumber, DateTime docDate)
        {
            File.Copy(templatePath, outputFileName, true);

            using (WordprocessingDocument doc = WordprocessingDocument.Open(outputFileName, true))
            {
                var body = doc.MainDocumentPart.Document.Body;

                var allTextElements = body.Descendants<Text>();

                foreach (var text in allTextElements)
                {
                    text.Text = text.Text
                        .Replace("count", count.ToString())
                        .Replace("decnumber", decNumber)
                        .Replace("docnumber", docNumber.ToString())
                        .Replace("docdate", docDate.ToString("dd.MM.yyyy"));
                }

                /*
                var textIterator = allTextElements.GetEnumerator();

                while (textIterator.MoveNext())
                {
                    var currentText = textIterator.Current;
                    currentText.Text = currentText.Text
                        .Replace("[count]", count.ToString())
                        .Replace("[dec_number]", decNumber.ToString())
                        .Replace("[doc_number]", docNumber.ToString())
                        .Replace("[doc_date]", docDate.ToString("dd.MM.yyyy"));
                }
                */
            }
        }
    }
}
