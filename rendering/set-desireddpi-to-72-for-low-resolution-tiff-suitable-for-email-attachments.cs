using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple document with a few pages.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("First page of the document.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Second page of the document.");

        // Configure image save options for TIFF with low resolution (72 DPI).
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // DesiredDpi is represented by the Resolution property in current API.
            Resolution = 72f
        };

        // Define output path in the current working directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "LowResolution.tiff");

        // Save the document as a multipage TIFF.
        doc.Save(outputPath, saveOptions);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The TIFF file was not created.", outputPath);
    }
}
