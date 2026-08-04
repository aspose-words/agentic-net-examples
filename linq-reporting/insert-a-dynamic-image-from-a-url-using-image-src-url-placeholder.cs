using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words (required for some encodings)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data model with image bytes from an embedded Base64 PNG
        ReportModel model = new()
        {
            ImageData = Convert.FromBase64String(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X9WcAAAAASUVORK5CYII=")
        };

        // Create a template document programmatically
        string templatePath = "Template.docx";
        CreateTemplate(templatePath);

        // Load the template
        Document doc = new(templatePath);

        // Build the report using LINQ Reporting Engine
        ReportingEngine engine = new();
        bool success = engine.BuildReport(doc, model, "model");

        // Save the generated report
        string outputPath = "Report.docx";
        doc.Save(outputPath);

        // Indicate completion (no interactive input)
        Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}. Output saved to '{outputPath}'.");
    }

    private static void CreateTemplate(string path)
    {
        Document template = new();
        DocumentBuilder builder = new(template);

        // Insert a textbox to host the image tag
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);
        // Image tag referencing the ImageData property of the model
        builder.Write("<<image [ImageData] -fitSize>>");

        // Save the template
        template.Save(path);
    }
}

// Data model used by the report
public class ReportModel
{
    // Image data to be inserted into the report (byte array)
    public byte[] ImageData { get; set; } = Array.Empty<byte>();
}
