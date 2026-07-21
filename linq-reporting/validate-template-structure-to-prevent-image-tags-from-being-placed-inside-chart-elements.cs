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
        const string templatePath = "Template.docx";
        const string outputPath = "Report.docx";
        const string imagePath = "sample.png";

        // Create a minimal image file used by the image tag.
        CreateSampleImage(imagePath);

        // -----------------------------------------------------------------
        // Step 1: Build the template document programmatically.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert a chart into the template.
        builder.InsertChart(ChartType.Column, 400, 300);

        // Retrieve the chart shape and (incorrectly) set a title that contains an image tag.
        Shape chartShape = (Shape)template.GetChildNodes(NodeType.Shape, true)[0];
        chartShape.Chart.Title.Text = "<<image [ImagePath]>>"; // Invalid placement for demonstration.

        // Insert a valid image tag outside the chart.
        builder.Writeln("<<image [ImagePath]>>");

        // Save the template to disk.
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // Step 2: Load the template and validate its structure.
        // -----------------------------------------------------------------
        Document loadedTemplate = new Document(templatePath);

        if (!ValidateTemplate(loadedTemplate))
        {
            Console.WriteLine("Template validation failed: image tag found inside a chart element.");
            return;
        }

        // -----------------------------------------------------------------
        // Step 3: Prepare the data model.
        // -----------------------------------------------------------------
        ReportModel model = new ReportModel
        {
            ImagePath = imagePath
        };

        // -----------------------------------------------------------------
        // Step 4: Build the report using the LINQ Reporting engine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };
        engine.BuildReport(loadedTemplate, model, "model");

        // -----------------------------------------------------------------
        // Step 5: Save the generated report.
        // -----------------------------------------------------------------
        loadedTemplate.Save(outputPath);
        Console.WriteLine("Report generated successfully.");
    }

    // Validates that no image tags are placed inside chart titles.
    private static bool ValidateTemplate(Document doc)
    {
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
        {
            // If the shape contains a chart, inspect its title.
            if (shape.Chart != null && shape.Chart.Title != null)
            {
                string title = shape.Chart.Title.Text ?? string.Empty;
                if (title.Contains("<<image"))
                {
                    return false; // Invalid image tag detected inside a chart title.
                }
            }
        }
        return true; // No invalid image tags found.
    }

    // Creates a 1x1 transparent PNG file for the sample image.
    private static void CreateSampleImage(string path)
    {
        byte[] pngData = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BAQAE/wJ" +
            "Z6VYAAAAASUVORK5CYII=");
        File.WriteAllBytes(path, pngData);
    }
}

// Simple data model used by the reporting engine.
public class ReportModel
{
    // Path to the image referenced by the image tag.
    public string ImagePath { get; set; } = string.Empty;
}
