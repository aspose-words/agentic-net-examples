using System;
using System.IO;
using System.IO.Compression;
using System.Collections.Generic;
using System.Text;

public class OleExtractor
{
    public static void Main()
    {
        // Prepare paths
        string tempDir = Path.Combine(Path.GetTempPath(), "OleExtractorDemo");
        Directory.CreateDirectory(tempDir);
        string docPath = Path.Combine(tempDir, "sample.docx");
        string dbPath = Path.Combine(tempDir, "oleobjects.bin");

        // Create a sample .docx (zip) with an embedded OLE object
        CreateSampleDocumentWithOle(docPath);

        // Extract OLE objects and store them in a simple binary "database"
        int storedCount = ExtractOleObjectsAndStore(docPath, dbPath);

        Console.WriteLine($"Extracted and stored {storedCount} OLE object(s) into the binary database at:");
        Console.WriteLine(dbPath);
    }

    private static void CreateSampleDocumentWithOle(string docPath)
    {
        // Ensure any existing file is removed
        if (File.Exists(docPath))
            File.Delete(docPath);

        // Create a zip archive that mimics a .docx structure
        using (var zip = ZipFile.Open(docPath, ZipArchiveMode.Create))
        {
            // Add a dummy OLE object under the typical embeddings folder
            var entry = zip.CreateEntry("word/embeddings/object1.bin");
            using (var entryStream = entry.Open())
            {
                // Example binary content of an OLE object
                byte[] oleData = new byte[] { 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0A };
                entryStream.Write(oleData, 0, oleData.Length);
            }

            // Add minimal required parts for a valid .docx (optional for this demo)
            var contentTypes = zip.CreateEntry("[Content_Types].xml");
            using (var ctStream = new StreamWriter(contentTypes.Open()))
            {
                ctStream.Write(@"<?xml version=""1.0"" encoding=""UTF-8""?>
<Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types"">
    <Default Extension=""bin"" ContentType=""application/vnd.openxmlformats-officedocument.oleObject""/>
</Types>");
            }
        }
    }

    private static int ExtractOleObjectsAndStore(string docPath, string dbPath)
    {
        var extractedObjects = new List<(string FileName, byte[] Data)>();

        // Open the document as a zip archive
        using (var zip = ZipFile.OpenRead(docPath))
        {
            // Iterate through entries that represent OLE objects
            foreach (var entry in zip.Entries)
            {
                if (IsOleEntry(entry.FullName))
                {
                    using (var entryStream = entry.Open())
                    using (var ms = new MemoryStream())
                    {
                        entryStream.CopyTo(ms);
                        extractedObjects.Add((Path.GetFileName(entry.FullName), ms.ToArray()));
                    }
                }
            }
        }

        // Write extracted objects to a simple binary file acting as a "database"
        using (var fs = new FileStream(dbPath, FileMode.Create, FileAccess.Write, FileShare.None))
        using (var bw = new BinaryWriter(fs, Encoding.UTF8, leaveOpen: false))
        {
            // Write the number of records first
            bw.Write(extractedObjects.Count);

            foreach (var (fileName, data) in extractedObjects)
            {
                // Write file name length and bytes (UTF-8)
                byte[] nameBytes = Encoding.UTF8.GetBytes(fileName);
                bw.Write(nameBytes.Length);
                bw.Write(nameBytes);

                // Write data length and bytes
                bw.Write(data.Length);
                bw.Write(data);
            }
        }

        return extractedObjects.Count;
    }

    private static bool IsOleEntry(string entryName)
    {
        // Typical locations for embedded OLE objects in Office Open XML packages
        return entryName.StartsWith("word/embeddings/", StringComparison.OrdinalIgnoreCase) ||
               entryName.StartsWith("embeddings/", StringComparison.OrdinalIgnoreCase) ||
               entryName.StartsWith("xl/embeddings/", StringComparison.OrdinalIgnoreCase) ||
               entryName.StartsWith("ppt/embeddings/", StringComparison.OrdinalIgnoreCase);
    }
}
