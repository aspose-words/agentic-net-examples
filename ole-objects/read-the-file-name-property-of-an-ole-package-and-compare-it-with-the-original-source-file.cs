using System;
using System.IO;
using System.Text;

public class OlePackageDemo
{
    public static void Main()
    {
        // Create a temporary source file
        string tempDir = Path.GetTempPath();
        string sourceFileName = "source.txt";
        string sourcePath = Path.Combine(tempDir, sourceFileName);
        File.WriteAllText(sourcePath, "Sample content for OLE package demo.");

        // Build a simple Ole10Native stream that mimics an OLE package
        byte[] oleStream = BuildOle10NativeStream(sourcePath);

        // Write the stream to a package file (for demo purposes)
        string packagePath = Path.Combine(tempDir, "package.bin");
        File.WriteAllBytes(packagePath, oleStream);

        // Read the filename property from the package
        string extractedFileName = ExtractFileNameFromOle10Native(packagePath);

        // Compare with the original source file name
        bool isMatch = string.Equals(extractedFileName, sourceFileName, StringComparison.OrdinalIgnoreCase);
        Console.WriteLine(isMatch ? "Match" : "Mismatch");

        // Clean up temporary files
        try { File.Delete(sourcePath); } catch { }
        try { File.Delete(packagePath); } catch { }
    }

    private static byte[] BuildOle10NativeStream(string sourcePath)
    {
        // Ole10Native format (simplified):
        // DWORD totalSize
        // DWORD unknown (0)
        // DWORD fileNameLength (bytes, Unicode)
        // Unicode file name + null terminator
        // DWORD sourcePathLength (bytes, Unicode)
        // Unicode source path + null terminator
        // DWORD tempPathLength (bytes, Unicode) - set to 0
        // DWORD fileDataLength
        // file data bytes

        string fileName = Path.GetFileName(sourcePath);
        string sourcePathUnicode = sourcePath;
        byte[] fileNameBytes = Encoding.Unicode.GetBytes(fileName + "\0");
        byte[] sourcePathBytes = Encoding.Unicode.GetBytes(sourcePathUnicode + "\0");
        byte[] fileData = File.ReadAllBytes(sourcePath);
        int fileDataLength = fileData.Length;

        using (var ms = new MemoryStream())
        using (var bw = new BinaryWriter(ms, Encoding.Unicode, true))
        {
            // Placeholder for totalSize (will fill later)
            bw.Write(0u);
            bw.Write(0u); // unknown

            bw.Write((uint)fileNameBytes.Length);
            bw.Write(fileNameBytes);

            bw.Write((uint)sourcePathBytes.Length);
            bw.Write(sourcePathBytes);

            bw.Write(0u); // tempPathLength (no temp path)

            bw.Write((uint)fileDataLength);
            bw.Write(fileData);

            // Calculate totalSize (size of the rest of the structure)
            uint totalSize = (uint)(ms.Length - 4);
            ms.Position = 0;
            bw.Write(totalSize);

            return ms.ToArray();
        }
    }

    private static string ExtractFileNameFromOle10Native(string packagePath)
    {
        using (var fs = new FileStream(packagePath, FileMode.Open, FileAccess.Read))
        using (var br = new BinaryReader(fs, Encoding.Unicode))
        {
            // Read totalSize (ignore)
            br.ReadUInt32();

            // Read unknown
            br.ReadUInt32();

            // Read file name length (in bytes)
            uint fileNameLength = br.ReadUInt32();
            byte[] fileNameBytes = br.ReadBytes((int)fileNameLength);
            string fileName = Encoding.Unicode.GetString(fileNameBytes);
            // Trim the terminating null character if present
            int nullIndex = fileName.IndexOf('\0');
            if (nullIndex >= 0)
                fileName = fileName.Substring(0, nullIndex);

            return fileName;
        }
    }
}
