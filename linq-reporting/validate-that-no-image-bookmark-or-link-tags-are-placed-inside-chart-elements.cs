using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using Aspose.Words.Reporting;

public class ReportModel
{
    // No properties needed for this simple example.
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // -----------------------------------------------------------------
        // Step 1: Create a document template that contains a chart.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 400, 300);

        // Intentionally place a prohibited tag inside the chart title for validation demo.
        chartShape.Chart.Title.Text = "<<image [model.ImagePath]>>";

        // Save the template (optional, just to illustrate the file exists).
        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Step 2: Validate that no image, bookmark, or link tags are inside chart elements.
        // -----------------------------------------------------------------
        bool hasProhibitedTags = false;
        string[] prohibitedPrefixes = { "<<image", "<<bookmark", "<<link" };

        // Iterate over all shapes that contain charts.
        foreach (Shape shape in templateDoc.GetChildNodes(NodeType.Shape, true))
        {
            if (shape.HasChart)
            {
                // Check the chart title for prohibited tags.
                string titleText = shape.Chart.Title?.Text ?? string.Empty;
                if (!string.IsNullOrEmpty(titleText))
                {
                    foreach (string prefix in prohibitedPrefixes)
                    {
                        if (titleText.Contains(prefix, StringComparison.Ordinal))
                        {
                            hasProhibitedTags = true;
                            Console.WriteLine(
                                $"Prohibited tag \"{prefix}\" found in chart title: \"{titleText}\"");
                        }
                    }
                }

                // Additional checks (e.g., legend, axis titles) could be added here similarly.
            }
        }

        if (!hasProhibitedTags)
        {
            Console.WriteLine("Validation passed: no prohibited tags inside chart elements.");
        }

        // -----------------------------------------------------------------
        // Step 3: Generate a clean report (chart without prohibited tags) using ReportingEngine.
        // -----------------------------------------------------------------
        Document cleanTemplate = new Document();
        DocumentBuilder cleanBuilder = new DocumentBuilder(cleanTemplate);

        Shape cleanChart = cleanBuilder.InsertChart(ChartType.Column, 400, 300);
        cleanChart.Chart.Title.Text = "Sales Overview";

        const string cleanTemplatePath = "CleanTemplate.docx";
        cleanTemplate.Save(cleanTemplatePath);

        // Load the clean template and build the report.
        Document reportDoc = new Document(cleanTemplatePath);
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };
        engine.BuildReport(reportDoc, new ReportModel(), "model");

        const string reportPath = "Report.docx";
        reportDoc.Save(reportPath);
        Console.WriteLine($"Report generated successfully: {reportPath}");
    }
}
