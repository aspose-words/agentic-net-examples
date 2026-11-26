using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a temporary working folder.
        string workFolder = Path.Combine(Path.GetTempPath(), "AsposeWordsBatchTiff");
        Directory.CreateDirectory(workFolder);

        // Create sample DOCX files.
        for (int i = 1; i <= 2; i++)
        {
            string docPath = Path.Combine(workFolder, $"Sample{i}.docx");
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"This is sample document {i}.");
            doc.Save(docPath);
        }

        // Process each DOCX file in the folder.
        string[] docxFiles = Directory.GetFiles(workFolder, "*.docx");
        foreach (string docxFile in docxFiles)
        {
            // Load the source document.
            Document document = new Document(docxFile);

            // Configure TIFF rendering options: 200 DPI and CCITT3 compression.
            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
            {
                // Resolution sets both DpiX and DpiY.
                Resolution = 200,
                TiffCompression = TiffCompression.Ccitt3
            };

            // Determine output TIFF path.
            string tiffPath = Path.ChangeExtension(docxFile, ".tiff");

            // Save the document as a TIFF image.
            document.Save(tiffPath, saveOptions);

            // Verify that the TIFF file was created.
            if (!File.Exists(tiffPath))
                throw new InvalidOperationException($"Failed to create TIFF file: {tiffPath}");
        }

        // Indicate completion.
        Console.WriteLine("Batch processing completed successfully.");
    }
}
