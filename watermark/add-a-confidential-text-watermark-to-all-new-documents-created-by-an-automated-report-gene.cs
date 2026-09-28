using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add sample content to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Automated report content goes here.");

        // Apply a confidential text watermark.
        doc.Watermark.SetText("CONFIDENTIAL");

        // Define the output file path.
        string outputPath = "ReportWithWatermark.docx";

        // Save the document with the watermark.
        doc.Save(outputPath);

        // Validate that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved successfully: {outputPath}");
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
