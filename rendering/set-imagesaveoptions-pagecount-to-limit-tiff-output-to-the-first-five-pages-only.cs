using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder to add content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add content that spans more than five pages.
        for (int i = 1; i <= 7; i++)
        {
            builder.Writeln($"This is page {i}.");
            // Insert a page break after each page except the last one.
            if (i < 7)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Configure image save options for a multi‑page TIFF.
        // Use PageSet to limit the output to the first five pages (zero‑based index).
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            PageSet = new PageSet(0, 5) // start at page 0, render 5 pages
        };

        string outputPath = "output.tiff";

        // Save the document as a TIFF file.
        doc.Save(outputPath, saveOptions);

        // Verify that the TIFF file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The TIFF file was not created.", outputPath);

        // Ensure the file is not empty.
        if (new FileInfo(outputPath).Length == 0)
            throw new Exception("The TIFF file is empty.");

        // Program completed successfully.
    }
}
