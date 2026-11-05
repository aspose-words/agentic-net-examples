using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Paths for the sample input and output documents.
        string inputPath = "SampleWithWatermark.docx";
        string outputPath = "SampleWithoutWatermark.docx";

        // -----------------------------------------------------------------
        // Create a sample document and add a text watermark.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document containing a watermark.");

        // Apply a text watermark using the Document.Watermark API.
        doc.Watermark.SetText("Sample Watermark");

        // Save the document that contains the watermark.
        doc.Save(inputPath);

        // -----------------------------------------------------------------
        // Load the document and remove all existing watermarks.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(inputPath);
        loadedDoc.Watermark.Remove();

        // Save the document after watermark removal.
        loadedDoc.Save(outputPath);

        // Simple validation that the output file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Watermark removed successfully. Output saved to '{outputPath}'.");
        }
        else
        {
            Console.WriteLine("Failed to save the output document.");
        }
    }
}
