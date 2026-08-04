using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for Aspose.Words in some environments)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare working directories
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample image file (1x1 pixel PNG)
        string imagePath = Path.Combine(outputDir, "sample.png");
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XcZcAAAAASUVORK5CYII=");
        File.WriteAllBytes(imagePath, pngBytes);

        // Create the template document
        string templatePath = Path.Combine(outputDir, "Template.docx");
        CreateTemplate(templatePath, imagePath);

        // Validate the template structure
        List<string> validationErrors = ValidateTemplate(templatePath);
        if (validationErrors.Count > 0)
        {
            Console.WriteLine("Template validation errors:");
            foreach (var err in validationErrors)
                Console.WriteLine("- " + err);
        }
        else
        {
            Console.WriteLine("Template validation passed. No image tags inside chart elements.");
        }

        // Build the report if validation passed
        if (validationErrors.Count == 0)
        {
            string reportPath = Path.Combine(outputDir, "Report.docx");
            BuildReport(templatePath, reportPath, imagePath);
            Console.WriteLine($"Report generated: {reportPath}");
        }
    }

    private static void CreateTemplate(string templatePath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a chart shape (correctly without any image tags inside)
        Shape chartShape = builder.InsertChart(ChartType.Column, 400, 300);
        // Optionally add a caption or placeholder after the chart
        builder.Writeln("Chart placeholder");

        // Insert a textbox shape with a correct image tag
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 200);
        builder.MoveTo(textBox.FirstParagraph);
        builder.Writeln("<<image [model.ImagePath] -fitSize>>");

        // Save the template
        doc.Save(templatePath);
    }

    private static List<string> ValidateTemplate(string templatePath)
    {
        List<string> errors = new();
        Document doc = new Document(templatePath);
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

        // Regex to detect image tags
        Regex imageTagRegex = new(@"<<\s*image\s*\[.*?\]\s*(?:-[\w]+)*\s*>>", RegexOptions.IgnoreCase);

        foreach (Shape shape in shapes)
        {
            if (shape.HasChart)
            {
                // Scan all paragraphs inside the chart shape for image tags
                foreach (Paragraph para in shape.GetChildNodes(NodeType.Paragraph, true))
                {
                    if (imageTagRegex.IsMatch(para.GetText()))
                    {
                        errors.Add($"Image tag found inside chart shape at paragraph text \"{para.GetText().Trim()}\".");
                    }
                }
            }
        }

        return errors;
    }

    private static void BuildReport(string templatePath, string reportPath, string imagePath)
    {
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };

        ReportModel model = new ReportModel
        {
            ImagePath = imagePath // Use full path so the engine can locate the image
        };

        engine.BuildReport(reportDoc, model, "model");
        reportDoc.Save(reportPath);
    }
}

public class ReportModel
{
    // ImagePath can be a relative or absolute path; using absolute path for simplicity
    public string ImagePath { get; set; } = string.Empty;
}
