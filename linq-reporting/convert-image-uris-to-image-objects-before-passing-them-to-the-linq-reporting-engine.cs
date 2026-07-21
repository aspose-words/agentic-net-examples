using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare a sample image file.
        string imagePath = Path.Combine(Directory.GetCurrentDirectory(), "sample.png");
        CreateSamplePng(imagePath);

        // Build the data model and load the image as a byte array.
        ReportModel model = new()
        {
            Title = "Image URI to Image Object Demo",
            ImageUri = imagePath,
            ImageData = File.ReadAllBytes(imagePath)
        };

        // Create a Word document template programmatically.
        Document doc = new();
        DocumentBuilder builder = new(doc);

        // Insert a title.
        builder.Writeln("<<[model.Title]>>");
        builder.Writeln();

        // Insert a textbox that will contain the image.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 300, 200);
        builder.MoveTo(textBox.FirstParagraph);
        builder.Write("<<image [model.ImageData] -fitSize>>");

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ReportOutput.docx");
        doc.Save(outputPath);
    }

    private static void CreateSamplePng(string path)
    {
        // Minimal 1x1 pixel PNG (transparent).
        byte[] pngBytes = new byte[]
        {
            0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A,
            0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52,
            0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,
            0x08,0x06,0x00,0x00,0x00,0x1F,0x15,0xC4,
            0x89,0x00,0x00,0x00,0x0A,0x49,0x44,0x41,
            0x54,0x78,0x9C,0x63,0x60,0x00,0x00,0x00,
            0x02,0x00,0x01,0xE2,0x21,0xBC,0x33,0x00,
            0x00,0x00,0x00,0x49,0x45,0x4E,0x44,0xAE,
            0x42,0x60,0x82
        };
        File.WriteAllBytes(path, pngBytes);
    }
}

public class ReportModel
{
    public string Title { get; set; } = "";
    public string ImageUri { get; set; } = "";
    public byte[] ImageData { get; set; } = Array.Empty<byte>();
}
