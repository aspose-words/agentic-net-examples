using System;
using System.IO;
using System.IO.Compression;
using System.Collections.Generic;

public class Program
{
    public static void Main()
    {
        // Path to the .docx file (adjust as needed)
        string docxPath = "sample.docx";

        // Output directory for extracted OLE objects
        string outputDir = "ExportedOleObjects";

        if (!File.Exists(docxPath))
        {
            // No document to process; exit gracefully
            return;
        }

        Directory.CreateDirectory(outputDir);

        using (FileStream fs = new FileStream(docxPath, FileMode.Open, FileAccess.Read, FileShare.Read))
        using (ZipArchive archive = new ZipArchive(fs, ZipArchiveMode.Read))
        {
            // OLE objects are stored under word/embeddings/
            const string embeddingsFolder = "word/embeddings/";
            int objectCounter = 1;

            foreach (ZipArchiveEntry entry in archive.Entries)
            {
                if (!entry.FullName.StartsWith(embeddingsFolder, StringComparison.OrdinalIgnoreCase))
                    continue;

                using (Stream entryStream = entry.Open())
                using (MemoryStream ms = new MemoryStream())
                {
                    entryStream.CopyTo(ms);
                    byte[] data = ms.ToArray();

                    string extension = DetectExtension(data);
                    string fileName = $"Object{objectCounter}{extension}";
                    string outputPath = Path.Combine(outputDir, fileName);

                    File.WriteAllBytes(outputPath, data);
                    objectCounter++;
                }
            }
        }
    }

    // Simple magic‑number based detection for common file types
    private static string DetectExtension(byte[] data)
    {
        if (data.Length >= 4)
        {
            // PDF: %PDF
            if (data[0] == 0x25 && data[1] == 0x50 && data[2] == 0x44 && data[3] == 0x46)
                return ".pdf";

            // PNG: 89 50 4E 47
            if (data[0] == 0x89 && data[1] == 0x50 && data[2] == 0x4E && data[3] == 0x47)
                return ".png";

            // JPEG: FF D8 FF
            if (data[0] == 0xFF && data[1] == 0xD8 && data[2] == 0xFF)
                return ".jpg";

            // DOC (old binary): D0 CF 11 E0
            if (data[0] == 0xD0 && data[1] == 0xCF && data[2] == 0x11 && data[3] == 0xE0)
                return ".doc";

            // XLS (old binary): D0 CF 11 E0 (same as DOC) – use .xls as fallback
            // We'll differentiate by checking later bytes if needed, but for simplicity return .xls
            // OLE compound files can be many types; default to .bin if unknown
        }

        if (data.Length >= 2)
        {
            // ZIP (e.g., .xlsx, .pptx, .docx) – start with PK
            if (data[0] == 0x50 && data[1] == 0x4B)
                return ".zip";
        }

        // Default fallback
        return ".bin";
    }
}
