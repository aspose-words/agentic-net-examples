using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare a temporary folder for the example files.
        string outputFolder = Path.Combine(Path.GetTempPath(), "AsposeWordsExample");
        Directory.CreateDirectory(outputFolder);

        // Paths for the source DOCX and the resulting TIFF.
        string docPath = Path.Combine(outputFolder, "sample.docx");
        string tiffPath = Path.Combine(outputFolder, "firstPage.tiff");

        // Create a multi‑page DOCX document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is the content of the first page.");
        // Insert a page break to create a second page.
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is the content of the second page.");
        // Save the source document.
        doc.Save(docPath);

        // Load the document from the file system.
        Document loadedDoc = new Document(docPath);

        // Configure ImageSaveOptions to produce a single‑page TIFF of the first page.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);
        // PageSet uses zero‑based page index; 0 selects the first page.
        options.PageSet = new PageSet(0);

        // Render and save the first page as a TIFF image.
        loadedDoc.Save(tiffPath, options);

        // Verify that the TIFF file was created.
        if (!File.Exists(tiffPath))
        {
            throw new Exception("The TIFF file was not created as expected.");
        }

        // Optionally, output the location of the generated file.
        Console.WriteLine($"First page TIFF saved to: {tiffPath}");
    }
}
