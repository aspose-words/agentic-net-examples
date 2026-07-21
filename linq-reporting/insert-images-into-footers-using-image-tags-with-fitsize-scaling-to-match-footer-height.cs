using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;   // For ShapeType

public class Program
{
    public static void Main()
    {
        // Create a simple 1x1 red PNG image and save it locally.
        string imageFile = Path.Combine(Directory.GetCurrentDirectory(), "sample.png");
        byte[] pngData = new byte[]
        {
            0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A,0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52,
            0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,0x08,0x02,0x00,0x00,0x00,0x90,0x77,0x53,
            0xDE,0x00,0x00,0x00,0x0A,0x49,0x44,0x41,0x54,0x08,0xD7,0x63,0xF8,0xCF,0xC0,0x00,
            0x00,0x04,0x00,0x01,0xE2,0x26,0x05,0x9B,0x00,0x00,0x00,0x00,0x49,0x45,0x4E,0x44,
            0xAE,0x42,0x60,0x82
        };
        File.WriteAllBytes(imageFile, pngData);

        // Data model exposing the image path.
        var model = new ReportModel { ImagePath = imageFile };

        // Build the template document programmatically.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        CreateTemplate(templatePath);

        // Load the template, run the LINQ Reporting engine, and save the result.
        var doc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        doc.Save(outputPath);
    }

    // Creates a Word document with a footer that contains a textbox holding an image tag.
    private static void CreateTemplate(string filePath)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Move to the primary footer of the first section.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);

        // Insert a textbox that will host the image.
        var textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);

        // Write the LINQ Reporting image tag with -fitSize switch.
        builder.Write("<<image [model.ImagePath] -fitSize>>");

        // Return to the main body and add a placeholder paragraph.
        builder.MoveToDocumentEnd();
        builder.Writeln("Report generated with image in footer.");

        // Save the template.
        doc.Save(filePath);
    }
}

// Simple data model used by the LINQ Reporting engine.
public class ReportModel
{
    public string ImagePath { get; set; } = string.Empty;
}
