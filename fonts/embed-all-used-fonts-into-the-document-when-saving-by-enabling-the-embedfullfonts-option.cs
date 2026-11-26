using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set font properties and add some text.
        builder.Font.Name = "Arial";
        builder.Font.Size = 24;
        builder.Writeln("Hello, this text will have its font embedded in the PDF.");

        // Configure PDF save options to embed all used fonts.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            EmbedFullFonts = true
        };

        // Save the document.
        string outputPath = "EmbeddedFontsOutput.pdf";
        doc.Save(outputPath, saveOptions);

        // Verify that the file was created.
        bool fileExists = File.Exists(outputPath);
        Console.WriteLine($"PDF saved. File exists: {fileExists}");
    }
}
