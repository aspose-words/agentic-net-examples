using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Automated report content goes here.");

        // Apply a confidential text watermark to the document.
        doc.Watermark.SetText("CONFIDENTIAL");

        // Save the document with the watermark.
        string outputPath = "ReportWithConfidentialWatermark.docx";
        doc.Save(outputPath);

        // Simple validation that the file was created.
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
