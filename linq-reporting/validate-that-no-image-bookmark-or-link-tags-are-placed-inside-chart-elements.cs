using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Initialize to avoid nullable warnings.
    public string Title { get; set; } = "Sample Report";
}

public class Program
{
    public static void Main()
    {
        // Paths for the template and the final report.
        const string templatePath = "ChartTemplate.docx";
        const string reportPath = "ChartReport.docx";

        // -----------------------------------------------------------------
        // 1. Create a template document that contains a chart.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a LINQ Reporting tag.
        builder.Writeln("<<[model.Title]>>");

        // Insert a simple column chart.
        builder.InsertChart(ChartType.Column, 400, 300);

        // OPTIONAL: Uncomment the lines below to simulate an invalid tag inside the chart.
        // Shape chartShape = (Shape)templateDoc.GetChildNodes(NodeType.Shape, true)[0];
        // chartShape.FirstParagraph?.AppendChild(new Run(templateDoc, "<<image [model.ImagePath]>>"));

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template back (simulating a real-world scenario).
        // -----------------------------------------------------------------
        Document loadedTemplate = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Validate that no image, bookmark, or link tags exist inside chart elements.
        // -----------------------------------------------------------------
        ValidateChartTags(loadedTemplate);

        // -----------------------------------------------------------------
        // 4. Build the report using LINQ Reporting Engine.
        // -----------------------------------------------------------------
        var model = new ReportModel(); // Sample data source.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(loadedTemplate, model, "model");

        // -----------------------------------------------------------------
        // 5. Save the generated report.
        // -----------------------------------------------------------------
        loadedTemplate.Save(reportPath);
    }

    // Scans all chart shapes and ensures they do not contain prohibited tags.
    private static void ValidateChartTags(Document doc)
    {
        // Retrieve all Shape nodes that contain a chart.
        NodeCollection chartShapes = doc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in chartShapes)
        {
            if (!shape.HasChart) continue;

            // Search all descendant runs within the chart shape.
            NodeCollection runs = shape.GetChildNodes(NodeType.Run, true);
            foreach (Run run in runs)
            {
                string text = run.Text;
                if (text.Contains("<<image") ||
                    text.Contains("<<bookmark") ||
                    text.Contains("<<link"))
                {
                    // Invalid tag detected – raise an exception with details.
                    string message = $"Prohibited tag found inside a chart at run hash {run.GetHashCode()}.";
                    throw new InvalidOperationException(message);
                }
            }
        }
    }
}
