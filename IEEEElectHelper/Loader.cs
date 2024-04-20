using Newtonsoft.Json;
using OfficeOpenXml;

namespace IEEEElectHelper;

public class Loader
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
        // Example: "IEEE AESS ESPRIT SBC Vice Chair" => "AESS"

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

    private static bool IsMemberOfChapter(string ieeeID , string position)
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

    public static void Testagain()
    {
        bool ismember = true;
        string excelFilePath = @"C:\Users\xtige\RiderProjects\IEEEElectHelper\IEEEElectHelper\bin\Debug\net8.0\assets\test.xlsx";
        if (!File.Exists("assets\\results\\debug.txt"))
        {
            File.Create("assets\\results\\debug.txt").Dispose();
        }
        File.WriteAllText("assets\\results\\debug.txt", "");

        using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelWorksheet worksheet = package.Workbook.Worksheets["Réponses au formulaire 1"];
            int rowCount = worksheet.Dimension.Rows;



            Dictionary<string, Dictionary<string, int>> chapterResults = new Dictionary<string, Dictionary<string, int>>(); 

            for (int row = 2; row <= rowCount; row++)
            {
                string ieeeID = worksheet.Cells[row, 2].Value.ToString(); 


                for (int col = 3; col <= worksheet.Dimension.Columns; col++) 
                {
                    string position = worksheet.Cells[1, col].Value.ToString();
                    if (position.Contains("Are"))
                    {
                        continue;
                    }

                    string vote = worksheet.Cells[row, col].Value?.ToString(); 

                    if (!string.IsNullOrEmpty(vote))
                    {
                        if (!chapterResults.ContainsKey(position))
                        {
                            chapterResults[position] = new Dictionary<string, int>(); 
                        }

                        if (!chapterResults[position].ContainsKey(vote))
                        {
                            chapterResults[position][vote] = 0; 
                        }

                        if (IsMemberOfChapter(ieeeID, position))
                        {
                            chapterResults[position][vote]++;
                            string ph = $"{ieeeID} voted for {vote}  {position} \n";
                            File.AppendAllText("assets\\results\\debug.txt", ph );

                        }
                        else
                        {
                            string voteee = $"{position} , Skipped {ieeeID} Vote was {vote}";
                            Console.WriteLine($"Skipped {ieeeID} because is not a member / invalid , Vote was {vote}  {position}");
                                using (System.IO.StreamWriter file = new System.IO.StreamWriter("assets\\results\\skipped\\skipped.txt", true)) // 'true' to append to the file if it exists
                                {
                                    file.WriteLine($"Skipped vote details: {voteee}");
                                    file.WriteLine(); // Add an empty line for separation
                                }
                            }

                    }
                }
                File.AppendAllText("assets\\results\\debug.txt", "\n");

            }
            foreach (var kvp in chapterResults)
            {

                string position = kvp.Key;
                if (position.Contains("Are"))
                {
                    continue;
                }
                Dictionary<string, int> optionResults = kvp.Value;

                // Write to text file
                string fileName = $"assets\\results\\{position}_Results.txt";
                using (System.IO.StreamWriter file = new System.IO.StreamWriter(fileName))
                {
                    file.WriteLine($"Position: {position}");
                    foreach (var optionKvp in optionResults)
                    {
                        string option = optionKvp.Key;
                        int voteCount = optionKvp.Value;
                        file.WriteLine($"{option}: {voteCount}");
                    }
                }
            }

        }
    }
    public static void writeskipped(Dictionary<string, Dictionary<string, int>> chapterResults , string txt)
    {
        
    }

}