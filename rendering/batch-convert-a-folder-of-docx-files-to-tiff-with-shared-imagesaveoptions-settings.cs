using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define folders for source DOCX files and resulting TIFF files.
        string sourceFolder = Path.Combine(Directory.GetCurrentDirectory(), "SourceDocs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "TiffOutput");

        // Ensure the folders exist.
        Directory.CreateDirectory(sourceFolder);
        Directory.CreateDirectory(outputFolder);

        // Create a few sample DOCX files locally.
        for (int i = 1; i <= 3; i++)
        {
            string docPath = Path.Combine(sourceFolder, $"Sample{i}.docx");
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"This is sample document {i}.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln($"Page two of sample document {i}.");
            doc.Save(docPath);
        }

        // Configure shared ImageSaveOptions for TIFF conversion.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Example shared settings.
            Resolution = 300,                     // 300 DPI.
            TiffCompression = TiffCompression.Lzw // LZW compression.
        };

        // Process each DOCX file in the source folder.
        foreach (string docFile in Directory.GetFiles(sourceFolder, "*.docx"))
        {
            // Load the DOCX document.
            Document document = new Document(docFile);

            // Determine the output TIFF file path.
            string tiffFileName = Path.GetFileNameWithoutExtension(docFile) + ".tiff";
            string tiffPath = Path.Combine(outputFolder, tiffFileName);

            // Save the document as a multipage TIFF using the shared options.
            document.Save(tiffPath, tiffOptions);

            // Validate that the TIFF file was created.
            if (!File.Exists(tiffPath))
                throw new InvalidOperationException($"Failed to create TIFF file: {tiffPath}");
        }

        // Indicate successful completion.
        Console.WriteLine("Batch conversion completed successfully.");
    }
}
