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
        builder.Writeln("This is a sample document.");

        // Apply a text watermark to the document.
        doc.Watermark.SetText("CONFIDENTIAL");

        // Save the document as a DOCX file.
        string outputPath = "WatermarkedDocument.docx";
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved successfully: {outputPath}");
        }
    }
}
