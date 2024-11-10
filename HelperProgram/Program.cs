using System;
using System.IO;
using System.Text.RegularExpressions;
using OfficeOpenXml;
using OfficeOpenXml.Style;

class Program
{
    static void Main()
    {
        ReadFile();
    }
    static void ReadFile()
    {
        // Đường dẫn tới thư mục gốc chứa các file .txt và các thư mục con
        string rootFolderPath = @"D:\WorkPlace\c#\HelperProgram\HelperProgram\TextFiles";
        // Đường dẫn tới file Excel cần ghi các cụm từ tiếng Trung
        string outputFilePath = @"D:\WorkPlace\c#\HelperProgram\HelperProgram\Output.xlsx";

        try
        {
            string folderName = new DirectoryInfo(rootFolderPath).Name;
            // Sử dụng EPPlus để tạo file Excel và ghi dữ liệu
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage package = new ExcelPackage())
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(folderName);

                int row = 2;

                // Lấy danh sách tất cả các file .txt trong thư mục chính và các thư mục con
                string[] files = Directory.GetFiles(rootFolderPath, "*.txt", SearchOption.AllDirectories);
                foreach (string filePath in files)
                {
                    string fileName = Path.GetFileName(filePath);

                    // Đọc nội dung của file txt
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

    }
}
