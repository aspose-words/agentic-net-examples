using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to insert a run of text.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Write("Sample text with double wave underline.");

        // Retrieve the run that was just added.
        Run run = (Run)builder.CurrentParagraph.LastChild;

        // Set the underline type to double wave (WavyDouble).
        run.Font.Underline = Underline.WavyDouble;

        // Validate that the underline type was set correctly.
        if (run.Font.Underline != Underline.WavyDouble)
            throw new InvalidOperationException("Underline type was not set correctly.");

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DoubleWaveUnderline.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the output file exists.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output file was not created.", outputPath);
    }
}
