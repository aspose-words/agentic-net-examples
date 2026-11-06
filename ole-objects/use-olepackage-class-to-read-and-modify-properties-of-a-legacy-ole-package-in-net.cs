using System;
using System.IO;
using System.Text;

public class OlePackage
{
    public byte[] Data { get; set; }
    public string Label { get; set; }
    public string FileName { get; set; }
    public string Path { get; set; }

    public OlePackage(byte[] data, string label, string fileName, string path)
    {
        Data = data;
        Label = label;
        FileName = fileName;
        Path = path;
    }

    public void SaveToFile(string filePath)
    {
        using var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write);
        using var bw = new BinaryWriter(fs, Encoding.UTF8, leaveOpen: false);
        bw.Write(Label ?? string.Empty);
        bw.Write(FileName ?? string.Empty);
        bw.Write(Path ?? string.Empty);
        bw.Write(Data?.Length ?? 0);
        if (Data != null && Data.Length > 0)
        {
            bw.Write(Data);
        }
    }

    public static OlePackage FromFile(string filePath)
    {
        using var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);
        using var br = new BinaryReader(fs, Encoding.UTF8, leaveOpen: false);
        string label = br.ReadString();
        string fileName = br.ReadString();
        string path = br.ReadString();
        int dataLength = br.ReadInt32();
        byte[] data = br.ReadBytes(dataLength);
        return new OlePackage(data, label, fileName, path);
    }
}

public class OlePackageDemo
{
    public static void Main()
    {
        // Prepare sample data for the OLE package
        byte[] sampleData = Encoding.UTF8.GetBytes("Hello OLE Package");

        // Create an OlePackage instance with initial properties
        OlePackage package = new OlePackage(
            data: sampleData,
            label: "OriginalLabel",
            fileName: "example.txt",
            path: @"C:\Temp"
        );

        // Define file paths for saving
        string originalFile = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "package_original.bin");
        string modifiedFile = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "package_modified.bin");

        // Save the original package to disk
        package.SaveToFile(originalFile);
        Console.WriteLine($"Original package saved to: {originalFile}");

        // Load the package back from the file
        OlePackage loadedPackage = OlePackage.FromFile(originalFile);
        Console.WriteLine($"Loaded package label: {loadedPackage.Label}");

        // Modify a property (e.g., the label)
        loadedPackage.Label = "ModifiedLabel";

        // Save the modified package to a new file
        loadedPackage.SaveToFile(modifiedFile);
        Console.WriteLine($"Modified package saved to: {modifiedFile}");

        // Verify modification by loading again
        OlePackage verifiedPackage = OlePackage.FromFile(modifiedFile);
        Console.WriteLine($"Verified package label: {verifiedPackage.Label}");
    }
}
