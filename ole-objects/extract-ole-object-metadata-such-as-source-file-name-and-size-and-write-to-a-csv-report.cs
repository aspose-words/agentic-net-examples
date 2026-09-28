using System;
using System.IO;
using System.Text;

public class OleMetadataExtractor
{
    public static void Main()
    {
        // Create a temporary directory for demo OLE objects
        string tempDir = Path.Combine(Path.GetTempPath(), "OleObjectsDemo");
        Directory.CreateDirectory(tempDir);

        // Create some dummy OLE object files
        CreateDummyOleFile(Path.Combine(tempDir, "image1.ole"), 1024);
        CreateDummyOleFile(Path.Combine(tempDir, "document2.ole"), 2048);
        CreateDummyOleFile(Path.Combine(tempDir, "presentation3.ole"), 3072);

        // Prepare CSV report
        string csvPath = Path.Combine(tempDir, "OleMetadataReport.csv");
        var sb = new StringBuilder();
        sb.AppendLine("FileName,SizeBytes");

        foreach (string filePath in Directory.EnumerateFiles(tempDir, "*.ole"))
        {
            FileInfo info = new FileInfo(filePath);
            string fileName = info.Name;
            long size = info.Length;
            sb.AppendLine($"{EscapeCsv(fileName)},{size}");
        }

        File.WriteAllText(csvPath, sb.ToString(), Encoding.UTF8);
    }

    private static void CreateDummyOleFile(string path, int sizeInBytes)
    {
        byte[] data = new byte[sizeInBytes];
        new Random().NextBytes(data);
        File.WriteAllBytes(path, data);
    }

    private static string EscapeCsv(string field)
    {
        if (field.Contains(",") || field.Contains("\"") || field.Contains("\n"))
        {
            return $"\"{field.Replace("\"", "\"\"")}\"";
        }
        return field;
    }
}
