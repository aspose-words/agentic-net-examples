using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class DocxToTiffBatch
{
    public static void Main()
    {
        // Define a temporary working folder.
        string workFolder = Path.Combine(Path.GetTempPath(), "DocxToTiffDemo");
        Directory.CreateDirectory(workFolder);

        // Create sample DOCX files in the folder.
        CreateSampleDocx(Path.Combine(workFolder, "Sample1.docx"),
            "First sample document.\nLine two.");
        CreateSampleDocx(Path.Combine(workFolder, "Sample2.docx"),
            "Second sample document.\nAnother line.");

        // Convert each DOCX file in the folder to a TIFF image using predefined compression.
        foreach (string docxPath in Directory.GetFiles(workFolder, "*.docx"))
        {
            // Load the DOCX document.
            Document doc = new Document(docxPath);

            // Prepare TIFF save options with a specific compression (CCITT4).
            ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
            {
                // Use the correct property name for TIFF compression.
                TiffCompression = TiffCompression.Ccitt4
            };

            // Determine the output TIFF file path.
            string tiffPath = Path.ChangeExtension(docxPath, ".tiff");

            // Save the document as a TIFF image.
            doc.Save(tiffPath, tiffOptions);

            // Validate that the TIFF file was created.
            if (!File.Exists(tiffPath))
                throw new InvalidOperationException($"Failed to create TIFF file: {tiffPath}");
        }

        // Indicate completion (no user interaction required).
        Console.WriteLine("DOCX to TIFF conversion completed successfully.");
    }

    // Helper method to create a simple DOCX file with given content.
    private static void CreateSampleDocx(string filePath, string content)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln(content);
        doc.Save(filePath);
    }
}
