using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare output folder and file name.
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputFolder);
        string pdfFile = Path.Combine(outputFolder, "WatermarkedDocument.pdf");

        // Create a new blank Word document.
        Document doc = new Document();

        // Add some content so the document has visible pages.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document with a watermark.");
        builder.Writeln("The watermark should be visible in the resulting PDF.");

        // Define watermark appearance.
        TextWatermarkOptions watermarkOptions = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 48,
            Color = Color.Gray,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = false
        };

        // Apply a text watermark to the document.
        doc.Watermark.SetText("CONFIDENTIAL", watermarkOptions);

        // Save the document directly as PDF, preserving the watermark.
        doc.Save(pdfFile, SaveFormat.Pdf);
    }
}
