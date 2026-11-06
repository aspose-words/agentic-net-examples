using System;
using System.IO;
using System.IO.Compression;
using System.Linq;

public class Program
{
    public static void Main()
    {
        string sourcePath = "sample.docx";
        string outputPath = "extracted_ole.bin";

        if (!File.Exists(sourcePath))
        {
            Console.WriteLine($"Source file '{sourcePath}' not found.");
            return;
        }

        using (FileStream fs = new FileStream(sourcePath, FileMode.Open, FileAccess.Read))
        using (ZipArchive archive = new ZipArchive(fs, ZipArchiveMode.Read))
        {
            ZipArchiveEntry oleEntry = archive.Entries
                .FirstOrDefault(e => e.FullName.StartsWith("word/embeddings/", StringComparison.OrdinalIgnoreCase));

            if (oleEntry == null)
            {
                Console.WriteLine("No OLE object found in the document.");
                return;
            }

            using (Stream entryStream = oleEntry.Open())
            using (FileStream outStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
            {
                entryStream.CopyTo(outStream);
            }

            Console.WriteLine($"OLE object extracted to '{outputPath}'.");
        }
    }
}
