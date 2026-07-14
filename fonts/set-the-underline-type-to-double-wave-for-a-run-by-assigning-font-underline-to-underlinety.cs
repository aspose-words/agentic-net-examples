using System;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the underline type to double wave (WavyDouble) for the current font.
        builder.Font.Underline = Underline.WavyDouble;

        // Write a line of text that will have the double‑wave underline.
        builder.Writeln("This text has a double‑wave underline.");

        // Validate that the underline was set correctly.
        if (builder.Font.Underline != Underline.WavyDouble)
            throw new InvalidOperationException("Underline type was not set to double wave.");

        // Save the document to the local file system.
        string outputPath = "DoubleWaveUnderline.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!System.IO.File.Exists(outputPath))
            throw new InvalidOperationException("The document was not saved correctly.");
    }
}
