using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define input and output folders relative to the current directory.
        string baseDir = Directory.GetCurrentDirectory();
        string inputFolder = Path.Combine(baseDir, "InputDocs");
        string outputFolder = Path.Combine(baseDir, "OutputTiffs");

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // -----------------------------------------------------------------
        // Bootstrap sample DOCX files if the input folder is empty.
        // -----------------------------------------------------------------
        if (Directory.GetFiles(inputFolder, "*.docx").Length == 0)
        {
            for (int i = 1; i <= 2; i++)
            {
                Document sampleDoc = new Document();
                DocumentBuilder builder = new DocumentBuilder(sampleDoc);
                builder.Writeln($"Sample document {i}");
                builder.InsertParagraph();
                builder.Writeln("This is a test paragraph to generate some content.");
                string samplePath = Path.Combine(inputFolder, $"Sample{i}.docx");
                sampleDoc.Save(samplePath);
            }
        }

        // -----------------------------------------------------------------
        // Process each DOCX file: render to a multi‑page TIFF at 200 DPI
        // using CCITT3 compression.
        // -----------------------------------------------------------------
        foreach (string docxPath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            // Load the source document.
            Document doc = new Document(docxPath);

            // Configure image save options for TIFF output.
            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
            {
                Resolution = 200,                     // 200 DPI.
                TiffCompression = TiffCompression.Ccitt3 // CCITT3 compression.
            };

            // Determine the output TIFF file name.
            string tiffFileName = Path.GetFileNameWithoutExtension(docxPath) + ".tiff";
            string tiffPath = Path.Combine(outputFolder, tiffFileName);

            // Save the document as a TIFF image.
            doc.Save(tiffPath, saveOptions);

            // Validate that the TIFF file was created.
            if (!File.Exists(tiffPath))
                throw new InvalidOperationException($"Failed to create TIFF file: {tiffPath}");
        }

        // Indicate successful completion (optional logging).
        Console.WriteLine("Batch processing completed successfully.");
    }
}
