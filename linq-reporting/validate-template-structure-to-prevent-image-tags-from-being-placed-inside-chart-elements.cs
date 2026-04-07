using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare a folder for all generated files.
        string workDir = Directory.GetCurrentDirectory();

        // -----------------------------------------------------------------
        // 1. Create a sample image file (1×1 pixel PNG).
        // -----------------------------------------------------------------
        string imagePath = Path.Combine(workDir, "sample.png");
        File.WriteAllBytes(imagePath, GetSamplePng());

        // -----------------------------------------------------------------
        // 2. Create the LINQ Reporting template and save it.
        // -----------------------------------------------------------------
        string templatePath = Path.Combine(workDir, "Template.docx");
        CreateTemplate(templatePath);

        // -----------------------------------------------------------------
        // 3. Validate that no <<image ...>> tags are placed inside chart shapes.
        // -----------------------------------------------------------------
        ValidateTemplate(templatePath);

        // -----------------------------------------------------------------
        // 4. Load the template for report generation.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 5. Prepare the data model.
        // -----------------------------------------------------------------
        ReportModel model = new ReportModel { ImagePath = imagePath };

        // -----------------------------------------------------------------
        // 6. Build the report using Aspose.Words LINQ Reporting engine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // -----------------------------------------------------------------
        // 7. Save the generated report.
        // -----------------------------------------------------------------
        string outputPath = Path.Combine(workDir, "Report.docx");
        doc.Save(outputPath);
    }

    // Creates a template containing a chart and an image tag placed inside a textbox (valid container).
    private static void CreateTemplate(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a chart shape – this will be the element we must keep free of image tags.
        builder.InsertChart(ChartType.Column, 400, 300);
        builder.Writeln(); // Ensure the cursor moves out of the chart.

        // Insert a textbox that will act as the image container.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        // Move the cursor inside the textbox's first paragraph.
        builder.MoveTo(textBox.FirstParagraph);
        // Write the image tag inside the textbox.
        builder.Write("<<image [model.ImagePath]>>");
        // Move the cursor back to the main document body after the textbox.
        builder.MoveToDocumentEnd();
        builder.Writeln(); // Optional line break.

        // Save the template to disk.
        doc.Save(filePath);
    }

    // Scans the document for any <<image ...>> tags that reside inside chart shapes.
    private static void ValidateTemplate(string filePath)
    {
        Document doc = new Document(filePath);
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
        {
            // Identify chart shapes.
            if (shape.HasChart)
            {
                // Retrieve any text that might be inside the chart shape.
                string innerText = shape.GetText();

                // If an image tag is found inside the chart, raise an exception.
                if (!string.IsNullOrEmpty(innerText) && innerText.Contains("<<image"))
                {
                    throw new InvalidOperationException("Image tags must not be placed inside chart elements.");
                }
            }
        }
    }

    // Minimal PNG (1×1 pixel, fully transparent) represented as a byte array.
    private static byte[] GetSamplePng()
    {
        return new byte[]
        {
            0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A,
            0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52,
            0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,
            0x08,0x06,0x00,0x00,0x00,0x1F,0x15,0xC4,
            0x89,0x00,0x00,0x00,0x0A,0x49,0x44,0x41,
            0x54,0x78,0x9C,0x63,0x60,0x00,0x00,0x00,
            0x02,0x00,0x01,0xE5,0x27,0xD4,0xA2,0x00,
            0x00,0x00,0x00,0x49,0x45,0x4E,0x44,0xAE,
            0x42,0x60,0x82
        };
    }
}

// Simple data model used by the LINQ Reporting engine.
public class ReportModel
{
    // Path to the image that will be inserted by the <<image>> tag.
    public string ImagePath { get; set; } = string.Empty;
}
