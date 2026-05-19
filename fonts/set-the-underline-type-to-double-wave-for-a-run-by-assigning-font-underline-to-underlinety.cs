using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "UnderlineDoubleWave.docx");

        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the underline type to double wave for the current font.
        builder.Font.Underline = Aspose.Words.Underline.WavyDouble;

        // Verify that the underline was set correctly.
        if (builder.Font.Underline != Aspose.Words.Underline.WavyDouble)
            throw new InvalidOperationException("Failed to set the underline type to double wave.");

        // Write a line of text that will display the double‑wave underline.
        builder.Writeln("This text has a double‑wave underline.");

        // Save the document to the specified file.
        doc.Save(outputPath);

        // Ensure that the file was created successfully.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The document was not saved correctly.", outputPath);
    }
}
