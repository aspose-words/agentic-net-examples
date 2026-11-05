using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a folder for sample DOCX files.
        string sourceFolder = Path.Combine(Directory.GetCurrentDirectory(), "SampleDocs");
        Directory.CreateDirectory(sourceFolder);

        // Generate a few sample DOCX documents.
        for (int i = 1; i <= 2; i++)
        {
            Document doc = new Document();
            Paragraph para = new Paragraph(doc);
            Run run = new Run(doc, $"This is sample document {i}.");
            para.AppendChild(run);
            doc.FirstSection.Body.AppendChild(para);

            string docPath = Path.Combine(sourceFolder, $"Doc{i}.docx");
            doc.Save(docPath);
        }

        // Create a folder for the TIFF output.
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "TiffOutput");
        Directory.CreateDirectory(outputFolder);

        // Define custom DPI for the TIFF images.
        const float customDpi = 300f;

        // Batch convert each DOCX file to a multipage TIFF.
        foreach (string docxPath in Directory.GetFiles(sourceFolder, "*.docx"))
        {
            Document document = new Document(docxPath);

            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
            {
                // Render all pages (default behavior) into a single multipage TIFF.
                // No need to set PageSet explicitly.
                Resolution = customDpi
            };

            string tiffPath = Path.Combine(outputFolder,
                Path.GetFileNameWithoutExtension(docxPath) + ".tiff");

            document.Save(tiffPath, saveOptions);

            // Verify that the TIFF file was created.
            if (!File.Exists(tiffPath))
                throw new InvalidOperationException($"Failed to create TIFF file: {tiffPath}");
        }

        // Indicate successful completion.
        Console.WriteLine("Batch conversion completed successfully.");
    }
}
