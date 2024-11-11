using System;
using System.IO;
using System.Text.RegularExpressions;
using OfficeOpenXml;
using OfficeOpenXml.Style;

class Program
{
    // Đường dẫn tới thư mục gốc chứa các file .txt và các thư mục con
    static string rootFolderPath = @"D:\WorkPlace\c#\HelperProgram\HelperProgram\TextFiles";
    // Đường dẫn tới file Excel cần ghi các cụm từ tiếng Trung
    static string outputFilePath = @"D:\WorkPlace\c#\HelperProgram\HelperProgram\Output.xlsx";
    static void Main()
    {
        ReadFile();
    }
    static void ReadFile()
    {
       
        try
        {
            string folderName = new DirectoryInfo(rootFolderPath).Name;
            // Sử dụng EPPlus để tạo file Excel và ghi dữ liệu
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage package = new ExcelPackage())
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(folderName);

                int row = 2;
                string[] files = Directory.GetFiles(rootFolderPath, "*.txt", SearchOption.AllDirectories);
                foreach (string filePath in files)
                {
                    string fileName = Path.GetFileName(filePath);
                    string text = File.ReadAllText(filePath);

                    // Biểu thức chính quy để tìm các cụm từ có ký tự Trung Quốc
                    string chinesePattern = @"([\u4e00-\u9fff]+(?:\s*[\u4e00-\u9fff]+)*)";

                    // Tìm tất cả các cụm từ chứa ký tự Trung Quốc
                    MatchCollection matches = Regex.Matches(text, chinesePattern);

                    if (matches.Count > 0)
                    {
                        // Ghi tên file vào cột A
                        worksheet.Cells[row, 1].Value = fileName;
                        worksheet.Cells[row, 1].Style.Font.Bold = true;
                        row++;

                        // Ghi các cụm từ tiếng Trung vào cột B
                        foreach (Match match in matches)
                        {
                            worksheet.Cells[row, 2].Value = match.Value.Trim();
                            row++;
                        }

                        // Thêm một hàng trống sau mỗi file để dễ đọc hơn
                        row++;
                    }
                    else
                    {
                        // Nếu không có cụm từ tiếng Trung nào, chỉ ghi tên file
                        worksheet.Cells[row, 1].Value = fileName;
                        worksheet.Cells[row, 1].Style.Font.Bold = true;
                        row++;

                        // Thêm một hàng trống
                        row++;
                    }
                }

                // Định dạng cột để vừa nội dung
                worksheet.Column(1).AutoFit();
                worksheet.Column(2).AutoFit();

                // Lưu file Excel
                FileInfo excelFile = new FileInfo(outputFilePath);
                package.SaveAs(excelFile);
            }

            Console.WriteLine("Success");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Fail: " + ex.Message);
        }
    }
    static void WriteFile()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        // Dictionary để lưu trữ cặp từ Trung-Việt
        Dictionary<string, string> translations = new Dictionary<string, string>();

        // Danh sách để lưu tên các file
        List<string> filenames = new List<string>();

        // Đọc file Excel
        using (var package = new ExcelPackage(new FileInfo(outputFilePath)))
        {
            var worksheet = package.Workbook.Worksheets[0];
            int rowCount = worksheet.Dimension.Rows;

            // Đọc dữ liệu từ Excel và lưu vào các collection
            for (int row = 2; row <= rowCount; row++)
            {
                string filename = worksheet.Cells[row, 1].Text?.Trim();
                string chinese = worksheet.Cells[row, 2].Text?.Trim();
                string vietnamese = worksheet.Cells[row, 3].Text?.Trim();

                if (!string.IsNullOrEmpty(filename))
                {
                    filenames.Add(filename);
                }

                if (!string.IsNullOrEmpty(chinese) && !string.IsNullOrEmpty(vietnamese))
                {
                    translations[chinese] = vietnamese;
                }
            }
        }

        // Xử lý từng file
        foreach (string filename in filenames)
        {
            if (File.Exists(filename))
            {
                try
                {
                    // Đọc nội dung file
                    string content = File.ReadAllText(filename);
                    string newContent = content;

                    // Thay thế các từ Trung bằng từ Việt
                    foreach (var translation in translations)
                    {
                        newContent = newContent.Replace(translation.Key, translation.Value);
                    }

                    // Tạo bản sao lưu của file gốc
                    string fileSaoLuu = Path.ChangeExtension(filename, ".backup" + Path.GetExtension(filename));
                    File.Copy(filename, fileSaoLuu, true);

                    // Ghi nội dung đã sửa vào file
                    File.WriteAllText(filename, newContent);

                    Console.WriteLine($"Success file: {filename}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error file {filename}: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine($"Not Found: {filename}");
            }
        }

        Console.WriteLine("Success");
    }

}
