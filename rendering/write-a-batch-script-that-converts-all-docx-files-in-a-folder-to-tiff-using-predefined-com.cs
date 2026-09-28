using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define a working folder for the demo.
        string workFolder = Path.Combine(Path.GetTempPath(), "DocxToTiffDemo");
        Directory.CreateDirectory(workFolder);

        // Create sample DOCX files in the folder.
        for (int i = 1; i <= 3; i++)
        {
            string docxPath = Path.Combine(workFolder, $"Sample{i}.docx");
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln($"This is sample document {i}.");
            // Add a second page to demonstrate multi‑page TIFF.
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln($"Second page of sample document {i}.");
            sampleDoc.Save(docxPath);
        }

        // Convert each DOCX file in the folder to a TIFF image with LZW compression.
        foreach (string docxFile in Directory.GetFiles(workFolder, "*.docx"))
        {
            // Load the source document.
            Document doc = new Document(docxFile);

            // Configure TIFF save options.
            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
            {
                // Use LZW compression for the TIFF output.
                TiffCompression = TiffCompression.Lzw
            };

            // Determine the output TIFF file path.
            string tiffFile = Path.ChangeExtension(docxFile, ".tiff");

            // Save the document as a TIFF image.
            doc.Save(tiffFile, saveOptions);

            // Validate that the TIFF file was created.
            if (!File.Exists(tiffFile))
                throw new InvalidOperationException($"Failed to create TIFF file: {tiffFile}");

            // Optional: ensure the file is not empty.
            if (new FileInfo(tiffFile).Length == 0)
                throw new InvalidOperationException($"TIFF file is empty: {tiffFile}");
        }

        // Indicate successful completion.
        Console.WriteLine("All DOCX files have been converted to TIFF successfully.");
    }
}
