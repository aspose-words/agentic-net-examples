using System;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add some sample content so the watermark is visible.
        var builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document.");
        builder.Writeln("The text watermark will appear behind this text.");

        // Apply a text watermark to the whole document.
        // Using the simple SetText method with default options.
        doc.Watermark.SetText("CONFIDENTIAL");

        // Optionally, customize the watermark appearance.
        var options = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 48,
            Color = System.Drawing.Color.Gray,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = true
        };
        doc.Watermark.SetText("CONFIDENTIAL", options);

        // Define the output file path (in the current working directory).
        string outputPath = "WatermarkedDocument.docx";

        // Save the document as DOCX.
        doc.Save(outputPath, SaveFormat.Docx);
    }
}
