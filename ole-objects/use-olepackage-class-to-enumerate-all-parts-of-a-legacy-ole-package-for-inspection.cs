using System;
using System.IO;
using System.Text;

public static class DataFormats
{
    public const string FileDrop = "FileDrop";
}

public class OlePackage
{
    private readonly string[] _files;

    private OlePackage(string[] files)
    {
        _files = files;
    }

    // Simulate creation from a single file; in a real OLE package this could contain many files.
    public static OlePackage FromFile(string filePath)
    {
        if (!File.Exists(filePath))
            throw new FileNotFoundException("File not found.", filePath);

        return new OlePackage(new[] { filePath });
    }

    public bool GetDataPresent(string format)
    {
        return format == DataFormats.FileDrop && _files.Length > 0;
    }

    public object GetData(string format)
    {
        if (format == DataFormats.FileDrop)
            return _files;
        return null;
    }
}

public class Program
{
    public static void Main()
    {
        // Create a temporary file to embed in the OLE package
        string tempFilePath = Path.Combine(Path.GetTempPath(), "SampleOlePackage.txt");
        File.WriteAllText(tempFilePath, "This is sample content inside an OLE package.");

        // Create an OlePackage from the temporary file
        OlePackage olePackage = OlePackage.FromFile(tempFilePath);

        // Enumerate the parts contained in the package (FileDrop format holds the file names)
        if (olePackage.GetDataPresent(DataFormats.FileDrop))
        {
            string[] files = (string[])olePackage.GetData(DataFormats.FileDrop);
            Console.WriteLine($"OLE package contains {files.Length} file(s):");
            foreach (string file in files)
            {
                Console.WriteLine($"- {Path.GetFileName(file)}");
                byte[] data = File.ReadAllBytes(file);
                Console.WriteLine($"  Size: {data.Length} bytes");
                Console.WriteLine($"  Content: {Encoding.UTF8.GetString(data)}");
            }
        }
        else
        {
            Console.WriteLine("No files found in the OLE package.");
        }

        // Clean up the temporary file
        try { File.Delete(tempFilePath); } catch { }
    }
}
