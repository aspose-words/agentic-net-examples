using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add a simple paragraph to demonstrate content.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Automated report generated on " + DateTime.Now);

        // Apply a confidential text watermark to the entire document.
        doc.Watermark.SetText("CONFIDENTIAL");

        // Define output path.
        string outputPath = "ReportWithWatermark.docx";

        // Save the document.
        doc.Save(outputPath);

        // Simple validation: ensure the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine("Document saved successfully with watermark: " + outputPath);
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
