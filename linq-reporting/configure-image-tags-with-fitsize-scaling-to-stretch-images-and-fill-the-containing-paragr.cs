using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

public class Program
{
    // Simple data model used by the LINQ Reporting template.
    public class ReportModel
    {
        // Path to the image that will be inserted into the report.
        public string ImagePath { get; set; } = string.Empty;
    }

    public static void Main()
    {
        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a tiny PNG image (1x1 pixel) and save it locally.
        // -----------------------------------------------------------------
        string imageFile = Path.Combine(outputDir, "sample.png");
        // Base64 representation of a 1x1 transparent PNG.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BAQAE/wJ/lZLZAAAAAElFTkSuQmCC";
        byte[] imageBytes = Convert.FromBase64String(base64Png);
        File.WriteAllBytes(imageFile, imageBytes);

        // -----------------------------------------------------------------
        // 2. Build the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert a textbox that will act as the image container.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 300, 200);
        // Move the cursor inside the textbox.
        builder.MoveTo(textBox.FirstParagraph);
        // Insert the image tag with the -fitSize switch to stretch the image.
        builder.Write("<<image [model.ImagePath] -fitSize>>");

        // -----------------------------------------------------------------
        // 3. Prepare the data source.
        // -----------------------------------------------------------------
        ReportModel model = new ReportModel
        {
            ImagePath = imageFile
        };

        // -----------------------------------------------------------------
        // 4. Build the report using ReportingEngine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        // The root object name must match the tag prefix used in the template ("model").
        engine.BuildReport(template, model, "model");

        // -----------------------------------------------------------------
        // 5. Save the generated document.
        // -----------------------------------------------------------------
        string resultPath = Path.Combine(outputDir, "ReportWithFitSizeImage.docx");
        template.Save(resultPath);
    }
}
