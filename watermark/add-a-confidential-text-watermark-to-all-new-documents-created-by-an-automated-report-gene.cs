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

        // Add sample content to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Automated Report");
        builder.Writeln("Generated on " + DateTime.Now);

        // Define watermark appearance.
        TextWatermarkOptions watermarkOptions = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 48,
            Color = Color.Red,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = false
        };

        // Apply the "CONFIDENTIAL" text watermark.
        doc.Watermark.SetText("CONFIDENTIAL", watermarkOptions);

        // Save the watermarked document.
        string outputFile = Path.Combine(Directory.GetCurrentDirectory(), "ConfidentialReport.docx");
        doc.Save(outputFile);
    }
}
