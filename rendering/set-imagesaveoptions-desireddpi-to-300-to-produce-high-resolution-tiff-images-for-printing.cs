using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple document with one paragraph.
        Document doc = new Document();
        Paragraph para = new Paragraph(doc);
        Run run = new Run(doc, "Hello World!");
        para.AppendChild(run);
        doc.FirstSection.Body.AppendChild(para);

        // Configure image save options for TIFF with high resolution (300 DPI).
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Set the resolution (DPI) for the output image.
            Resolution = 300
        };

        // Define output file path.
        string outputPath = "output.tiff";

        // Save the document as a TIFF image using the specified options.
        doc.Save(outputPath, saveOptions);

        // Verify that the TIFF file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the TIFF file at '{outputPath}'.");
        }

        // Confirmation message.
        Console.WriteLine($"TIFF image saved successfully at '{Path.GetFullPath(outputPath)}' with 300 DPI.");
    }
}
