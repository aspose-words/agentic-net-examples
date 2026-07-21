using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define folders for input DOCX files and output TIFF files.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputTiffs");

        // Ensure clean environment.
        if (Directory.Exists(inputFolder)) Directory.Delete(inputFolder, true);
        if (Directory.Exists(outputFolder)) Directory.Delete(outputFolder, true);
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample DOCX files.
        CreateSampleDocument(Path.Combine(inputFolder, "Sample1.docx"), "First document", 2);
        CreateSampleDocument(Path.Combine(inputFolder, "Sample2.docx"), "Second document", 3);

        // Set desired DPI for the TIFF output.
        const float dpi = 300f;

        // Prepare save options for multipage TIFF.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            Resolution = dpi,
            PageLayout = MultiPageLayout.TiffFrames()
        };

        // Process each DOCX file in the input folder.
        foreach (string docxPath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            // Load the DOCX document.
            Document doc = new Document(docxPath);

            // Determine output TIFF file name.
            string tiffFileName = Path.GetFileNameWithoutExtension(docxPath) + ".tiff";
            string tiffPath = Path.Combine(outputFolder, tiffFileName);

            // Save the document as a multipage TIFF.
            doc.Save(tiffPath, tiffOptions);

            // Verify that the TIFF file was created.
            if (!File.Exists(tiffPath))
                throw new InvalidOperationException($"Failed to create TIFF file: {tiffPath}");
        }

        // All conversions completed successfully.
        Console.WriteLine("Batch conversion completed. TIFF files are located at:");
        Console.WriteLine(outputFolder);
    }

    // Helper method to create a simple DOCX document with a given title and number of pages.
    private static void CreateSampleDocument(string filePath, string title, int pageCount)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Font.Name = "Times New Roman";
        builder.Font.Size = 24;
        builder.Writeln(title);
        builder.Writeln($"Generated on {DateTime.Now}");

        for (int i = 1; i < pageCount; i++)
        {
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln($"Page {i + 1}");
        }

        doc.Save(filePath, SaveFormat.Docx);
    }
}
