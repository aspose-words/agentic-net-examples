using System;
using System.IO;
using System.Text;

public class OlePackage
{
    public string FileName { get; }
    public string DisplayName { get; }
    public byte[] Data { get; }

    public OlePackage(byte[] data, string fileName, string tempFileName, string displayName)
    {
        Data = data ?? Array.Empty<byte>();
        FileName = fileName ?? string.Empty;
        DisplayName = displayName ?? string.Empty;
        // tempFileName is kept for compatibility with the original signature but not used here.
    }
}

public class Program
{
    public static void Main()
    {
        // Sample binary content for the OLE package
        byte[] originalData = Encoding.UTF8.GetBytes("Hello OLE Package");
        string fileName = "sample.txt";
        string tempFileName = Path.GetTempFileName();
        string displayName = "Sample OLE";

        // Create the OlePackage instance
        OlePackage package = new OlePackage(originalData, fileName, tempFileName, displayName);

        // Display original properties
        Console.WriteLine($"Original FileName: {package.FileName}");
        Console.WriteLine($"Original DisplayName: {package.DisplayName}");
        Console.WriteLine($"Original Data: {Encoding.UTF8.GetString(package.Data)}");

        // Modify the data (create a new package with updated content)
        byte[] modifiedData = Encoding.UTF8.GetBytes("Modified content");
        OlePackage modifiedPackage = new OlePackage(modifiedData, package.FileName, tempFileName, package.DisplayName);

        // Write the modified data to a file
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "modified.bin");
        File.WriteAllBytes(outputPath, modifiedPackage.Data);

        Console.WriteLine($"Modified data written to: {outputPath}");
    }
}
