using OfficeOpenXml;
namespace ExcelFileHandling
{
    public class FileModel
    {
        public int SrNo { get; set; }
        public string? SerialNo { get; set; }
        public string? PartDescription { get; set; }
        public string? Supplier { get; set; }
    }

    class Program
    {
    static void Main()
        {
            ExcelPackage.License.SetNonCommercialPersonal("My Name");

            string filePath = @"C:\Users\VishV\Desktop\OORJA\ExcelFileHandling\Sample.xlsx";
            Console.WriteLine("-------Welcome To the Excel Handling Program-------");
            MainMenu(filePath);

            Console.WriteLine("\nPress A to restart the program or B to exit.");
            string? finalChoice = Console.ReadLine();

            if (finalChoice?.ToUpper() == "A")
            {
                Main(); // Restart program
            }
            else
            {
                Console.WriteLine("Exiting...");
            }
        }

        static void MainMenu(string file)
        {
            bool run = true;
            FileInfo fileInfo = new FileInfo(file);

            while (run)
            {
                Console.WriteLine("\nOptions:");
                Console.WriteLine("1. Display Columns");
                Console.WriteLine("2. Search Part by SerialNo");
                Console.WriteLine("3. Add New Part");
                Console.WriteLine("4. View a Part");
                Console.WriteLine("5. Delete a Part");
                Console.WriteLine("6. View Entire Excel File");
                Console.WriteLine("7. Delete Entire Excel File");
                Console.WriteLine("8. Exit Menu");

                Console.Write("Enter your choice: ");
                string? inp = Console.ReadLine();

                switch (inp)
                {
                    case "1": DisplayColumns(fileInfo); break;
                    case "2": SearchPart(fileInfo); break;
                    case "3": AddNewPart(fileInfo); break;
                    case "4": ViewPart(fileInfo); break;
                    case "5": DeletePart(fileInfo); break;
                    case "6": ViewAllParts(fileInfo); break;
                    case "7": DeleteExcelFile(fileInfo); break;
                    case "8": Console.WriteLine("Exiting menu..."); run = false; break;
                    default: Console.WriteLine("Invalid choice. Please enter a number between 1-8."); break;
                }
            }
        }

        static List<FileModel> LoadParts(FileInfo file)
        {
            using var package = new ExcelPackage(file);
            var parts = new List<FileModel>();
            if (!file.Exists) return parts;
            var ws = package.Workbook.Worksheets.FirstOrDefault();
            if (ws == null) return parts;

            int rows = ws.Dimension?.Rows ?? 0;

            for (int i = 2; i <= rows; i++)
            {
                parts.Add(new FileModel
                {
                    SrNo = int.Parse(ws.Cells[i, 1].Text),
                    SerialNo = ws.Cells[i, 2].Text,
                    PartDescription = ws.Cells[i, 3].Text,
                    Supplier = ws.Cells[i, 4].Text
                });
            }

            return parts;
        }

        static void SaveParts(FileInfo file, List<FileModel> parts)
        {
            using var package = new ExcelPackage(file);
            var ws = package.Workbook.Worksheets.FirstOrDefault() ?? package.Workbook.Worksheets.Add("Parts");

            ws.Cells[1, 1].Value = "Sr No.";
            ws.Cells[1, 2].Value = "Serial Number";
            ws.Cells[1, 3].Value = "Part Description";
            ws.Cells[1, 4].Value = "Supplier";

            int row = 2;
            foreach (var p in parts)
            {
                ws.Cells[row, 1].Value = p.SrNo;
                ws.Cells[row, 2].Value = p.SerialNo?.ToString();
                ws.Cells[row, 3].Value = p.PartDescription?.ToString();
                ws.Cells[row, 4].Value = p.Supplier?.ToString();
                row++;
            }
            package.Save();
        }

        static void DisplayColumns(FileInfo file)
        {
            if (!file.Exists)
            {
                Console.WriteLine("File not found.");
                return;
            }

            using var package = new ExcelPackage(file);
            var ws = package.Workbook.Worksheets.FirstOrDefault();
            if (ws == null)
            {
                Console.WriteLine("Worksheet not found.");
                return;
            }

            int cols = ws.Dimension?.Columns ?? 0;
            Console.WriteLine("\nColumns:");
            for (int c = 1; c <= cols; c++)
            {
                Console.WriteLine($"{c}. {ws.Cells[1, c].Text}");
            }
        }

        static void SearchPart(FileInfo file)
        {
            var parts = LoadParts(file);
            if (parts.Count == 0)
            {
                Console.WriteLine("No parts to search.");
                return;
            }

            Console.Write("Enter SerialNo to search: ");
            string? input = Console.ReadLine()?.Trim() ?? "";

            var part = parts.FirstOrDefault(p => p.SerialNo?.Equals(input, StringComparison.OrdinalIgnoreCase) == true);
            if (part != null)
                DisplayPart(part);
            else
                Console.WriteLine("No part found with that SerialNo.");
        }

        static void AddNewPart(FileInfo file)
        {
            var parts = LoadParts(file);

            Console.Write("Enter Serial No.: ");
            string serialNo = Console.ReadLine()?.Trim() ?? "";

            if (parts.Any(p => p.SerialNo?.Equals(serialNo, StringComparison.OrdinalIgnoreCase) == true))
            {
                Console.WriteLine("This Serial No. already exists.");
                return;
            }

            Console.Write("Enter Part Description: ");
            string desc = Console.ReadLine() ?? "";

            Console.Write("Enter Supplier: ");
            string supplier = Console.ReadLine() ?? "";

            var newPart = new FileModel
            {
                SrNo = parts.Count + 1,
                SerialNo = serialNo,
                PartDescription = desc,
                Supplier = supplier
            };

            parts.Add(newPart);
            SaveParts(file, parts);

            Console.WriteLine("Part added.");
        }

        static void ViewPart(FileInfo file)
        {
            var parts = LoadParts(file);
            if (parts.Count == 0)
            {
                Console.WriteLine("No parts available.");
                return;
            }

            Console.Write("Enter SerialNo to view: ");
            string? input = Console.ReadLine()?.Trim() ?? "";

            var part = parts.FirstOrDefault(p => p.SerialNo?.Equals(input, StringComparison.OrdinalIgnoreCase) == true);
            if (part != null)
                DisplayPart(part);
            else
                Console.WriteLine("No part found.");
        }

        static void DeletePart(FileInfo file)
        {
            var parts = LoadParts(file);
            if (parts.Count == 0)
            {
                Console.WriteLine("No parts to delete.");
                return;
            }

            Console.Write("Enter SerialNo to delete: ");
            string? input = Console.ReadLine()?.Trim() ?? "";

            var partToDelete = parts.FirstOrDefault(p => p.SerialNo?.Equals(input, StringComparison.OrdinalIgnoreCase) == true);
            if (partToDelete == null)
            {
                Console.WriteLine("No matching part found.");
                return;
            }

            parts.Remove(partToDelete);

            // Reassign SrNo
            for (int i = 0; i < parts.Count; i++)
                parts[i].SrNo = i + 1;

            SaveParts(file, parts);
            Console.WriteLine("Part deleted.");
        }

        static void ViewAllParts(FileInfo file)
        {
            var parts = LoadParts(file);
            if (parts.Count == 0)
            {
                Console.WriteLine("No parts found.");
                return;
            }

            Console.WriteLine("\nAll parts:");
            foreach (var p in parts)
            {
                DisplayPart(p);
            }
        }

        static void DeleteExcelFile(FileInfo file)
        {
            if (file.Exists)
            {
                file.Delete();
                Console.WriteLine("Excel file deleted.");
            }
            else
            {
                Console.WriteLine("Excel file does not exist.");
            }
        }

        static void DisplayPart(FileModel p)
        {
            Console.WriteLine($"Sr.No: {p.SrNo}");
            Console.WriteLine($"Serial No.: {p.SerialNo}");
            Console.WriteLine($"Part Description: {p.PartDescription}");
            Console.WriteLine($"Supplier: {p.Supplier}");
            Console.WriteLine("-------------------------");
        }
    }
}
