using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required by Aspose.Words)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare output folder
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create a sample image file (1x1 pixel PNG)
        string imagePath = Path.Combine(outputDir, "sample.png");
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8Xw8AAukB9YV6ZV8AAAAASUVORK5CYII=";
        File.WriteAllBytes(imagePath, Convert.FromBase64String(base64Png));

        // Build the template document
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a chart with an invalid image tag inside its title (should be detected)
        Shape chartShape = builder.InsertChart(ChartType.Column, 400, 300);
        chartShape.Chart.Title.Text = "<<image [model.ImagePath]>>";

        // Insert a textbox with a correct image tag
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);
        builder.Write("<<image [model.ImagePath] -fitSize>>");

        // Save the template
        string templatePath = Path.Combine(outputDir, "template.docx");
        templateDoc.Save(templatePath);

        // Validate template: ensure no image tags are inside chart elements
        Document validationDoc = new Document(templatePath);
        bool invalidFound = false;
        foreach (Shape shape in validationDoc.GetChildNodes(NodeType.Shape, true))
        {
            if (shape.Chart != null && shape.Chart.Title?.Text != null)
            {
                string titleText = shape.Chart.Title.Text;
                if (Regex.IsMatch(titleText, @"<<\s*image\s*\[.*?\]"))
                {
                    Console.WriteLine("Invalid template: image tag found inside a chart title.");
                    invalidFound = true;
                    break;
                }
            }
        }

        if (invalidFound)
        {
            // Stop processing due to validation failure
            return;
        }

        // Prepare the data model
        ReportModel model = new()
        {
            ImagePath = imagePath
        };

        // Build the report
        ReportingEngine engine = new();
        engine.Options = ReportBuildOptions.None;
        bool success = engine.BuildReport(validationDoc, model, "model");

        // Save the generated report
        string reportPath = Path.Combine(outputDir, "report.docx");
        validationDoc.Save(reportPath);

        Console.WriteLine(success
            ? $"Report generated successfully at: {reportPath}"
            : "Report generation failed.");
    }
}

public class ReportModel
{
    public string ImagePath { get; set; } = string.Empty;
}
