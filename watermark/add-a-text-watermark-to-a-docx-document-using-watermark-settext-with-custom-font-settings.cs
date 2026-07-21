using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add a simple paragraph so the document has visible content.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample document with a custom text watermark.");

        // Configure the text watermark options.
        TextWatermarkOptions options = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 36,
            Color = Color.Blue,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = false
        };

        // Apply the text watermark to the document.
        doc.Watermark.SetText("Confidential", options);

        // Define the output file path.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "WatermarkedDocument.docx");

        // Save the document.
        doc.Save(outputPath);

        // Simple verification that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved successfully to: {outputPath}");
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
