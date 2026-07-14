using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some sample content spanning a few pages.
        builder.Writeln("This is page 1.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is page 2.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is page 3.");

        // Configure image save options for TIFF with low resolution (72 DPI).
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Set both horizontal and vertical resolution to 72 DPI.
            Resolution = 72f
        };

        // Define the output file path in the current working directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "LowResolution.tiff");

        // Save the document as a multipage TIFF using the configured options.
        doc.Save(outputPath, saveOptions);

        // Verify that the file was created successfully.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The TIFF file was not created.");

        // Optional: output a confirmation message.
        Console.WriteLine($"TIFF saved successfully at: {outputPath}");
    }
}
