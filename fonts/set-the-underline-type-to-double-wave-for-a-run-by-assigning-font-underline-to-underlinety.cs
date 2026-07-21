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

        // Insert a run of text.
        builder.Write("This text has a double wave underline.");

        // Set the underline type to double wave.
        // The correct enum value is Underline.WavyDouble.
        builder.Font.Underline = Underline.WavyDouble;

        // Validate that the underline was set correctly.
        if (builder.Font.Underline == Underline.WavyDouble)
        {
            Console.WriteLine("Underline type successfully set to DoubleWave.");
        }
        else
        {
            Console.WriteLine("Failed to set underline type.");
        }

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document.
        string outputPath = Path.Combine(outputDir, "DoubleWaveUnderline.docx");
        doc.Save(outputPath);

        // Confirm that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved successfully to: {outputPath}");
        }
        else
        {
            Console.WriteLine("Document was not saved.");
        }
    }
}
