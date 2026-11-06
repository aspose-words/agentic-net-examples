using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file paths.
        string docPath = "SampleDocument.docx";
        string tiffPath = "SampleDocument.tiff";

        // Create a new document and add some colored content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Size = 24;
        builder.Font.Color = System.Drawing.Color.DarkBlue;
        builder.Writeln("This is a sample DOCX document rendered to a color TIFF.");

        // Save the DOCX file (optional, just to show the source exists).
        doc.Save(docPath);

        // Configure ImageSaveOptions for TIFF with 24bppRgb (color) pixel format.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Optional: set compression to none to keep full color fidelity.
            TiffCompression = TiffCompression.None,
            // Set a reasonable resolution.
            Resolution = 300
        };

        // Render the document to a multi-page TIFF.
        doc.Save(tiffPath, saveOptions);

        // Validate that the TIFF file was created.
        if (!File.Exists(tiffPath))
        {
            throw new Exception($"TIFF file was not created at path: {tiffPath}");
        }

        // Output a simple confirmation (no user interaction required).
        Console.WriteLine($"Document successfully rendered to color TIFF: {tiffPath}");
    }
}
