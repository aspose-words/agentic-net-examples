using System;
using System.IO;
using System.Text;

public class OlePackageDemo
{
    public static void Main()
    {
        string tempFile = Path.Combine(Path.GetTempPath(), "OlePackageDemo.tmp");
        string originalFileName = "sample.txt";

        // 1. Create a minimal Ole10Native binary file containing the file name.
        CreateOlePackage(tempFile, originalFileName);

        // 2. Read the file name back from the binary file.
        string extractedFileName = ReadFileNameFromOlePackage(tempFile);

        // 3. Compare and output result.
        bool match = string.Equals(originalFileName, extractedFileName, StringComparison.Ordinal);
        Console.WriteLine($"Original:  {originalFileName}");
        Console.WriteLine($"Extracted: {extractedFileName}");
        Console.WriteLine($"Match: {match}");

        // Clean up
        try { File.Delete(tempFile); } catch { }
    }

    private static void CreateOlePackage(string filePath, string fileName)
    {
        // Build Ole10Native data (simplified: only file name, no source/temp paths, no file data)
        byte[] fileNameBytes = Encoding.Default.GetBytes(fileName);
        int nameLen = fileNameBytes.Length + 1; // include terminating null

        using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None))
        using (var bw = new BinaryWriter(fs, Encoding.Default, true))
        {
            bw.Write(nameLen);                     // DWORD: length of file name (including null)
            bw.Write(fileNameBytes);               // file name bytes
            bw.Write((byte)0);                     // null terminator
            bw.Write(0);                           // DWORD: length of source path (0)
            bw.Write(0);                           // DWORD: length of temporary path (0)
            // No file data follows in this minimal example
        }
    }

    private static string ReadFileNameFromOlePackage(string filePath)
    {
        using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read))
        using (var br = new BinaryReader(fs, Encoding.Default, true))
        {
            int nameLen = br.ReadInt32();               // includes null terminator
            byte[] nameBytes = br.ReadBytes(nameLen);   // read the name + null
            string extractedName = Encoding.Default.GetString(nameBytes).TrimEnd('\0');
            return extractedName;
        }
    }
}
