using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define source and destination folders.
        string sourceFolder = Path.Combine(Directory.GetCurrentDirectory(), "Docs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "Tiffs");

        // Ensure folders exist.
        Directory.CreateDirectory(sourceFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample DOCX files if none exist.
        if (Directory.GetFiles(sourceFolder, "*.docx").Length == 0)
        {
            for (int i = 1; i <= 2; i++)
            {
                Document sampleDoc = new Document();
                DocumentBuilder builder = new DocumentBuilder(sampleDoc);
                builder.Writeln($"Sample document {i}");
                builder.Writeln("This is a test paragraph.");
                string samplePath = Path.Combine(sourceFolder, $"Sample{i}.docx");
                sampleDoc.Save(samplePath);
            }
        }

        // Prepare TIFF conversion options with predefined compression.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            TiffCompression = TiffCompression.Lzw // Use LZW compression.
        };

        // Process each DOCX file in the source folder.
        foreach (string docxPath in Directory.GetFiles(sourceFolder, "*.docx"))
        {
            // Load the DOCX document.
            Document doc = new Document(docxPath);

            // Determine output TIFF file name.
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(docxPath);
            string tiffPath = Path.Combine(outputFolder, $"{fileNameWithoutExt}.tiff");

            // Save the document as a multi‑page TIFF using the options.
            doc.Save(tiffPath, tiffOptions);

            // Verify that the TIFF file was created.
            if (!File.Exists(tiffPath))
                throw new InvalidOperationException($"Failed to create TIFF file: {tiffPath}");
        }

        // Optional: indicate completion.
        Console.WriteLine("Conversion completed successfully.");
    }
}
