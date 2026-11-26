using System;
using System.IO;
using System.Text;

public class Program
{
    public static void Main(string[] args)
    {
        // Directory to scan – current directory
        string directory = Directory.GetCurrentDirectory();

        // Define OLE file extensions to look for
        string[] oleExtensions = new[] { ".ole", ".bin", ".doc", ".xls", ".ppt" };

        // Prepare CSV output
        string csvPath = Path.Combine(directory, "OleMetadataReport.csv");
        var sb = new StringBuilder();
        sb.AppendLine("FileName,SizeBytes");

        // Scan files
        foreach (string filePath in Directory.EnumerateFiles(directory, "*.*", SearchOption.AllDirectories))
        {
            string ext = Path.GetExtension(filePath);
            if (Array.Exists(oleExtensions, e => e.Equals(ext, StringComparison.OrdinalIgnoreCase)))
            {
                FileInfo info = new FileInfo(filePath);
                string line = $"{EscapeCsv(info.Name)},{info.Length}";
                sb.AppendLine(line);
            }
        }

        // Write CSV
        File.WriteAllText(csvPath, sb.ToString(), Encoding.UTF8);
    }

    // Simple CSV escaping for fields containing commas or quotes
    private static string EscapeCsv(string field)
    {
        if (field.Contains("\""))
            field = field.Replace("\"", "\"\"");
        if (field.Contains(",") || field.Contains("\"") || field.Contains("\n") || field.Contains("\r"))
            field = $"\"{field}\"";
        return field;
    }
}
