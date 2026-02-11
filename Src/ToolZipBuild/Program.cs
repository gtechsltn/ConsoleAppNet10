using System.IO.Compression;

namespace ToolZipBuild;

class Program
{
    static void Main()
    {
        Console.OutputEncoding = System.Text.Encoding.UTF8;

        const string publishDir = "D:\\gtechsltn\\ConsoleAppNet10\\Src\\ConsoleAppNet10\\bin\\Release\\net10.0\\win-x64\\publish";
        const string strNameSpace = "ConsoleAppNet10";
        string exeName = $"{strNameSpace}.exe";
        string zipName = $"{strNameSpace}.zip";
        string exePath = Path.Combine(publishDir, exeName);
        string zipPath = Path.Combine(publishDir, zipName);

        if (!File.Exists(exePath))
        {
            Console.WriteLine("Không tìm thấy file exe.");
            return;
        }

        // Xóa zip cũ nếu có
        if (File.Exists(zipPath))
            File.Delete(zipPath);

        using (ZipArchive zip = ZipFile.Open(zipPath, ZipArchiveMode.Create))
        {
            zip.CreateEntryFromFile(exePath, exeName, CompressionLevel.Optimal);
        }

        Console.WriteLine("Đã tạo file ZIP:");
        Console.WriteLine(zipPath);

        Console.Write("Nhấn phím bất kỳ để thoát ...");
        Console.ReadKey();
    }
}
