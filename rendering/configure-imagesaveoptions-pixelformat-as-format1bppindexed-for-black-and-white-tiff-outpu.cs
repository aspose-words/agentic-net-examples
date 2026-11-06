using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple document in memory.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello World!");

        // Save the source document (optional, just to demonstrate file creation).
        string sourcePath = "Sample.docx";
        doc.Save(sourcePath);

        // Configure ImageSaveOptions for black‑and‑white TIFF output.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff);
        tiffOptions.PixelFormat = ImagePixelFormat.Format1bppIndexed;

        // Render the document to a TIFF file using the configured options.
        string tiffPath = "Output.tiff";
        doc.Save(tiffPath, tiffOptions);

        // Verify that the TIFF file was created.
        if (!File.Exists(tiffPath))
            throw new FileNotFoundException("TIFF output was not created.", tiffPath);

        // Indicate successful completion.
        Console.WriteLine("Document successfully saved as black‑and‑white TIFF.");
    }
}
