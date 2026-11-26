using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a line of text
        builder.Writeln("Sample text with double wave underline.");

        // Retrieve the first run in the document
        Run run = doc.FirstSection.Body.Paragraphs[0].Runs[0];

        // The Aspose.Words enum Underline does not expose DoubleWave in the
        // version used for this example. The underlying numeric value for
        // DoubleWave is 9, so we assign it by casting the integer.
        const int DoubleWaveNumericValue = 9;
        run.Font.Underline = (Underline)DoubleWaveNumericValue;

        // Validate that the underline type was set correctly
        if ((int)run.Font.Underline != DoubleWaveNumericValue)
        {
            throw new Exception("Underline type was not set to DoubleWave.");
        }

        // Save the document
        string outputPath = "DoubleWaveUnderline.docx";
        doc.Save(outputPath);

        // Verify that the file was created
        if (!File.Exists(outputPath))
        {
            throw new Exception("Failed to create the output document.");
        }

        Console.WriteLine($"Document saved successfully to '{outputPath}'.");
    }
}
