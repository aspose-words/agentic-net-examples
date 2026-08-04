using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main(string[] args)
    {
        // Expect at least the directory path as the first argument.
        if (args.Length == 0)
            return; // No directory supplied; exit silently.

        string targetDirectory = args[0];

        // Optional second argument specifies the watermark text; default if omitted.
        string watermarkText = args.Length > 1 ? args[1] : "Confidential";

        // Validate the directory exists.
        if (!Directory.Exists(targetDirectory))
            return; // Invalid directory; exit.

        // Process all Word documents in the directory (both .docx and .doc).
        string[] files = Directory.GetFiles(targetDirectory, "*.*", SearchOption.TopDirectoryOnly);
        foreach (string filePath in files)
        {
            string extension = Path.GetExtension(filePath).ToLowerInvariant();
            if (extension != ".docx" && extension != ".doc")
                continue; // Skip non‑Word files.

            // Load the document.
            Document doc = new Document(filePath);

            // Apply a text watermark using the native API.
            doc.Watermark.SetText(watermarkText);

            // Overwrite the original file with the watermarked version.
            doc.Save(filePath);
        }
    }
}
