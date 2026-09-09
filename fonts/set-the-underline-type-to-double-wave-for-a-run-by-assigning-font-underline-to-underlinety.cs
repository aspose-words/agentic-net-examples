using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set underline type to double wave (WavyDouble) for the current font.
        builder.Font.Underline = Aspose.Words.Underline.WavyDouble;

        // Write sample text.
        builder.Writeln("This text is underlined with a double wave.");

        // Validate that the underline was set correctly.
        if (builder.Font.Underline != Aspose.Words.Underline.WavyDouble)
            throw new InvalidOperationException("Underline type was not set correctly.");

        // Define output path.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "DoubleWaveUnderline.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The document was not saved.", outputPath);
    }
}
