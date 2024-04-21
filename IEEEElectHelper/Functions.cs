using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IEEEElectHelper
{
    internal class Functions
    {
        public static void Test()
        {
            string excelFilePath = @"C:\Users\xtige\RiderProjects\IEEEElectHelper\IEEEElectHelper\bin\Debug\net8.0\assets\Members.xlsx";


            using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                ExcelWorksheet worksheet = package.Workbook.Worksheets["Sheet1"];
                int rowCount = worksheet.Dimension.Rows;

                var chapters = new Dictionary<string, List<string>>();
                string currentChapter = null;

                for (int row = 1; row <= rowCount; row++)
                {
                    string cellValue = worksheet.Cells[row, 1].Value?.ToString();

                    if (!string.IsNullOrWhiteSpace(cellValue))
                    {
                        currentChapter = cellValue;
                        string shortchap = GetChapterName(currentChapter);

                        if (!chapters.ContainsKey(currentChapter))
                        {
                            chapters.Add(currentChapter, new List<string>());

                        }
                    }
                    else if (!string.IsNullOrWhiteSpace(currentChapter))
                    {
                        string ieeeId = worksheet.Cells[row, 2].Value?.ToString();
                        if (!string.IsNullOrWhiteSpace(ieeeId))
                        {
                            chapters[currentChapter].Add(ieeeId);
                        }
                    }
                }
                foreach (var chapter in chapters)
                {
                    string chapterFilePath = $"assets\\chaptersFIX\\{chapter}_Members.txt";
                    if (!File.Exists(chapterFilePath))
                    {
                        File.Create(chapterFilePath).Dispose();
                    }
                    File.WriteAllText(chapterFilePath, "");
                    File.AppendAllText(chapterFilePath, string.Join("\n", chapter.Key + "\n"));
                    File.AppendAllText(chapterFilePath, string.Join("\n", chapter.Value));

                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"Chapter: {chapter.Key}");
                    Console.ForegroundColor = ConsoleColor.White;
                    Console.WriteLine($"Members: {string.Join(", ", chapter.Value)}");
                }

                string jsonFilePath = "assets\\output.json";
                File.WriteAllText(jsonFilePath, JsonConvert.SerializeObject(chapters, Formatting.Indented));

                Console.WriteLine("Parsing completed. Output saved to output.json");
            }

        }
        public static void check()
        {
            Console.WriteLine("enter id");
            var output = Console.ReadLine();
            string directoryPath = @"C:\Users\xtige\RiderProjects\IEEEElectHelper\IEEEElectHelper\bin\Debug\net8.0\assets\chaptersFIX";
            string[] files = Directory.GetFiles(directoryPath, "*.txt");

            foreach (string file in files)
            {
                string[] lines = File.ReadAllLines(file);

                foreach (string line in lines)
                {
                    string firstLine = File.ReadLines(file).FirstOrDefault();

                    if (line.Contains(output))
                    {
                        Console.WriteLine($"{firstLine}");
                        break;
                    }
                }
            }

        }



        public static string GetChapterName(string position)
        {
            string[] parts = position.Split(' ');
            if (parts.Length >= 2 && parts[1] != "ESP")
            {
                return parts[1];
            }
            else
            {
                return "UnknownChapter";
            }
        }

        public static bool IsMemberOfChapter(string ieeeID, string position)
        {
            string chapter = GetChapterName(position);
            string chapterFilePath = $"assets\\chapters\\{chapter}.txt";
            if (!File.Exists(chapterFilePath))
            {
                return false;
            }

            string[] ids = File.ReadAllLines(chapterFilePath);
            foreach (string id in ids)
            {
                if (id == ieeeID)
                {
                    return true;
                }
            }
            return false;
        }

        public static string[] RemoveEmptyLines(string[] lines)
        {
            return lines
                .Where(line => !string.IsNullOrWhiteSpace(line))
                .ToArray();
        }


    }
}
