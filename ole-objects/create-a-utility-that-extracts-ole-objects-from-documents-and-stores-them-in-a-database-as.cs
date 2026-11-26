using System;
using System.IO;
using System.IO.Compression;
using System.Text;

public class OleExtractor
{
    public static void Main()
    {
        string docPath = "sample.docx";
        string dbPath = "OleObjects.db";

        // Ensure a sample .docx file with an OLE object exists
        if (!File.Exists(docPath))
        {
            CreateSampleDocxWithOle(docPath);
        }

        // Ensure the database file exists (create empty if needed)
        if (!File.Exists(dbPath))
        {
            using (File.Create(dbPath)) { }
        }

        // Extract OLE objects from the .docx and store them as BLOBs
        int extractedCount = ExtractOleObjectsToDatabase(docPath, dbPath);

        // Write a simple log file with the result (no console interaction)
        File.WriteAllText("ExtractionLog.txt", $"Extracted {extractedCount} OLE object(s) from '{docPath}' into '{dbPath}'.");
    }

    private static void CreateSampleDocxWithOle(string path)
    {
        using (FileStream fs = new FileStream(path, FileMode.Create))
        using (ZipArchive archive = new ZipArchive(fs, ZipArchiveMode.Create))
        {
            // Minimal required parts for a .docx (empty document.xml)
            var docEntry = archive.CreateEntry("word/document.xml");
            using (var writer = new StreamWriter(docEntry.Open(), Encoding.UTF8))
            {
                writer.Write(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""><w:body></w:body></w:document>");
            }

            // Add a fake OLE object in the embeddings folder
            var oleEntry = archive.CreateEntry("word/embeddings/object1.bin");
            using (var oleStream = oleEntry.Open())
            {
                byte[] sampleOleData = Encoding.UTF8.GetBytes("Sample OLE content");
                oleStream.Write(sampleOleData, 0, sampleOleData.Length);
            }
        }
    }

    private static int ExtractOleObjectsToDatabase(string docPath, string dbPath)
    {
        int count = 0;
        using (FileStream fs = new FileStream(docPath, FileMode.Open, FileAccess.Read))
        using (ZipArchive archive = new ZipArchive(fs, ZipArchiveMode.Read))
        {
            foreach (var entry in archive.Entries)
            {
                if (entry.FullName.StartsWith("word/embeddings/", StringComparison.OrdinalIgnoreCase))
                {
                    using (var entryStream = entry.Open())
                    using (var ms = new MemoryStream())
                    {
                        entryStream.CopyTo(ms);
                        byte[] data = ms.ToArray();
                        AppendOleObject(dbPath, entry.FullName, data);
                        count++;
                    }
                }
            }
        }
        return count;
    }

    private static void AppendOleObject(string dbPath, string objectName, byte[] data)
    {
        using (FileStream fs = new FileStream(dbPath, FileMode.Append, FileAccess.Write))
        using (BinaryWriter writer = new BinaryWriter(fs, Encoding.UTF8, leaveOpen: false))
        {
            byte[] nameBytes = Encoding.UTF8.GetBytes(objectName);
            writer.Write(nameBytes.Length);
            writer.Write(nameBytes);
            writer.Write(data.Length);
            writer.Write(data);
        }
    }
}
