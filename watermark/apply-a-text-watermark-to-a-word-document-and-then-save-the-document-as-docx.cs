using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add some sample content.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document.");

        // Apply a text watermark to the entire document.
        doc.Watermark.SetText("CONFIDENTIAL");

        // Define the output file path.
        string outputPath = "WatermarkedDocument.docx";

        // Save the document as DOCX.
        doc.Save(outputPath);

        // Validate that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved successfully: {outputPath}");
        }
    }
}
