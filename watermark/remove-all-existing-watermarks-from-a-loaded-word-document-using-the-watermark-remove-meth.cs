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
        // Step 1: Create a sample document and add a text watermark.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document with a watermark.");

        // Add a text watermark using the Document.Watermark API.
        doc.Watermark.SetText("CONFIDENTIAL");

        // Save the document that contains the watermark.
        doc.Save(inputPath);

        // -----------------------------------------------------------------
        // Step 2: Load the document that contains the watermark.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(inputPath);

        // -----------------------------------------------------------------
        // Step 3: Remove all existing watermarks.
        // -----------------------------------------------------------------
        loadedDoc.Watermark.Remove();

        // -----------------------------------------------------------------
        // Step 4: Save the document after watermark removal.
        // -----------------------------------------------------------------
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
