using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class BatchDocxToTiff
{
    public static void Main()
    {
        // Define a temporary working folder.
        string workFolder = Path.Combine(Path.GetTempPath(), "DocxToTiffDemo");
        Directory.CreateDirectory(workFolder);

        // Create sample DOCX files in the folder.
        for (int i = 1; i <= 2; i++)
        {
            string docxPath = Path.Combine(workFolder, $"Sample{i}.docx");
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"This is sample document {i}.");
            doc.Save(docxPath);
        }

        // Prepare shared ImageSaveOptions for TIFF output.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Example: set resolution to 300 DPI.
            Resolution = 300
            // Additional options can be set here if needed.
        };

        // Process each DOCX file in the folder.
        foreach (string docxFile in Directory.GetFiles(workFolder, "*.docx"))
        {
            // Load the DOCX document.
            Document document = new Document(docxFile);

            // Determine output TIFF file path.
            string tiffFile = Path.ChangeExtension(docxFile, ".tiff");

            // Save the document as a multipage TIFF using the shared options.
            document.Save(tiffFile, tiffOptions);

            // Validate that the TIFF file was created.
            if (!File.Exists(tiffFile))
                throw new InvalidOperationException($"Failed to create TIFF file: {tiffFile}");
        }

        // Optional: indicate successful completion.
        Console.WriteLine("Batch conversion completed successfully.");
    }
}
