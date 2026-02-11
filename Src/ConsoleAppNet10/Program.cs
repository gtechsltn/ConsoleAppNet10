using ClosedXML.Excel;

class Program
{
    static void Main()
    {
        Console.OutputEncoding = System.Text.Encoding.UTF8;

        Console.WriteLine("QUẢN LÝ DANH MỤC ĐẦU TƯ CHUNG CƯ\n");

        string fileName = "DanhMucChungCu.xlsx";
        XLWorkbook wb;
        IXLWorksheet ws;

        // ===== LOAD / CREATE FILE =====
        if (File.Exists(fileName))
        {
            wb = new XLWorkbook(fileName);
            ws = wb.Worksheet("DanhMucChungCu");
        }
        else
        {
            wb = new XLWorkbook();
            ws = wb.Worksheets.Add("DanhMucChungCu");

            ws.Cell(1, 1).Value = "Căn hộ";
            ws.Cell(1, 2).Value = "Giá mua (tỷ)";
            ws.Cell(1, 3).Value = "Thuê/tháng (triệu)";
            ws.Cell(1, 4).Value = "Vay (%)";
            ws.Cell(1, 5).Value = "Lãi vay (%)";
            ws.Cell(1, 6).Value = "Dòng tiền/năm (triệu)";
            ws.Cell(1, 7).Value = "Yield thực (%)";
            ws.Cell(1, 8).Value = "Đánh giá";

            ws.Range(1, 1, 1, 8).Style.Font.Bold = true;
        }

        int row = ws.LastRowUsed()?.RowNumber() + 1 ?? 2;

        // ===== INPUT =====
        string name = ReadString("Tên căn hộ: ");
        double price = ReadDouble("Giá mua (tỷ): ", 2.5);
        double rentMonth = ReadDouble("Thuê/tháng (triệu): ", 12);
        double loanRate = ReadDouble("Tỷ lệ vay (%): ", 50) / 100;
        double loanInterest = ReadDouble("Lãi vay (%/năm): ", 9) / 100;
        double bankRate = ReadDouble("Lãi NH (%/năm): ", 6) / 100;
        double costRate = ReadDouble("Chi phí vận hành (%): ", 10) / 100;

        // ===== CALC =====
        double annualRent = rentMonth * 12 * (1 - costRate);
        double loanAmount = price * loanRate;
        double interestCost = loanAmount * loanInterest;
        double cashFlow = annualRent - interestCost;
        double yieldReal = annualRent / price;

        // ===== EVALUATE =====
        string rating;
        XLColor color;

        if (cashFlow >= 0 && yieldReal >= bankRate)
        {
            rating = "Tốt";
            color = XLColor.LightGreen;
        }
        else if (cashFlow > -30)
        {
            rating = "Cân nhắc";
            color = XLColor.LightYellow;
        }
        else
        {
            rating = "Không nên";
            color = XLColor.LightPink;
        }

        // ===== WRITE EXCEL =====
        ws.Cell(row, 1).Value = name;
        ws.Cell(row, 2).Value = price;
        ws.Cell(row, 3).Value = rentMonth;
        ws.Cell(row, 4).Value = loanRate * 100;
        ws.Cell(row, 5).Value = loanInterest * 100;
        ws.Cell(row, 6).Value = Math.Round(cashFlow, 1);
        ws.Cell(row, 7).Value = Math.Round(yieldReal * 100, 2);
        ws.Cell(row, 8).Value = rating;

        ws.Range(row, 1, row, 8).Style.Fill.BackgroundColor = color;

        ws.Columns().AdjustToContents();

        // Freeze dòng đầu (cố định dòng đầu tiên)
        ws.SheetView.FreezeRows(1);
        
        wb.SaveAs(fileName);

        Console.WriteLine($"\nĐã cập nhật file: {fileName}");
        
        Console.Write("Nhấn phím bất kỳ để thoát ...");
        Console.ReadKey();
    }

    // ===== HELPERS =====
    static double ReadDouble(string msg, double def)
    {
        Console.Write(msg);
        string? input = Console.ReadLine();
        if (!double.TryParse(input, out double val))
        {
            Console.WriteLine($"Dùng mặc định: {def}");
            return def;
        }
        return val;
    }

    static string ReadString(string msg)
    {
        Console.Write(msg);
        return Console.ReadLine() ?? "Không tên";
    }
}