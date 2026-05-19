using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    // Simple data model used as the root object for the report.
    public class ReportModel
    {
        public string Name { get; set; } = "Sample Name";
    }

    public static void Main()
    {
        // 1. Create a new blank document and a builder to compose the template.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // 2. Insert a chart into the document. This will be the element we need to validate.
        //    The chart is inserted as a Shape that has the HasChart flag set.
        Shape chartShape = builder.InsertChart(ChartType.Column, 400, 300);

        // 3. Add a regular LINQ Reporting tag outside the chart for demonstration purposes.
        builder.Writeln("<<[model.Name]>>");

        // 4. Save the template (optional, shown here for completeness).
        const string templatePath = "Template.docx";
        doc.Save(templatePath);

        // 5. Validate that no image, bookmark, or link tags are placed inside any chart element.
        bool tagsInsideChart = ContainsForbiddenTagsInsideCharts(doc);
        Console.WriteLine(tagsInsideChart
            ? "Validation failed: forbidden tags found inside a chart."
            : "Validation succeeded: no forbidden tags inside charts.");

        // 6. Prepare a simple data source for the LINQ Reporting engine.
        ReportModel model = new ReportModel();

        // 7. Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // 8. Save the final report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }

    // Scans the document for chart shapes and checks their descendant runs for forbidden tags.
    private static bool ContainsForbiddenTagsInsideCharts(Document document)
    {
        // Retrieve all Shape nodes in the document.
        NodeCollection shapes = document.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapes)
        {
            // Process only shapes that contain a chart.
            if (!shape.HasChart) continue;

            // Search for Run nodes inside the chart shape.
            NodeCollection runs = shape.GetChildNodes(NodeType.Run, true);
            foreach (Run run in runs)
            {
                string text = run.Text;
                // Check for any of the prohibited LINQ Reporting tags.
                if (text.Contains("<<image") ||
                    text.Contains("<<bookmark") ||
                    text.Contains("<<link"))
                {
                    return true; // Forbidden tag detected.
                }
            }
        }
        return false; // No forbidden tags found.
    }
}
