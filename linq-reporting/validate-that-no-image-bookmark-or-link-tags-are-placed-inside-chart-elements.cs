using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using Aspose.Words.Tables;

public class ReportModel
{
    // Root object for the template.
    public string Title { get; set; } = "Sales Report";

    // Collection of data items to be displayed in the table.
    public System.Collections.Generic.List<DataItem> Items { get; set; } = new()
    {
        new DataItem { Category = "Q1", Value = 120 },
        new DataItem { Category = "Q2", Value = 150 },
        new DataItem { Category = "Q3", Value = 180 },
        new DataItem { Category = "Q4", Value = 200 }
    };
}

public class DataItem
{
    public string Category { get; set; } = "";
    public int Value { get; set; }
}

public class Program
{
    private const string TemplatePath = "Template.docx";
    private const string OutputPath = "Report.docx";

    public static void Main()
    {
        // 1. Create a template document that contains a chart and LINQ Reporting tags.
        CreateTemplate();

        // 2. Load the template from disk.
        Document template = new Document(TemplatePath);

        // 3. Verify that no prohibited tags (image, bookmark, link) exist inside chart elements.
        ValidateNoProhibitedTagsInCharts(template);

        // 4. Build the report using the LINQ Reporting engine.
        var model = new ReportModel();
        var engine = new ReportingEngine { Options = ReportBuildOptions.None };
        engine.BuildReport(template, model, "model");

        // 5. Save the generated report.
        template.Save(OutputPath);
    }

    private static void CreateTemplate()
    {
        // Create a blank document and a builder to populate it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a chart. The chart itself must not contain any LINQ Reporting tags.
        Shape chartShape = builder.InsertChart(ChartType.Column, 400, 300);
        Chart chart = chartShape.Chart;
        chart.Title.Text = "Quarterly Sales";
        chart.Title.Show = true;

        // Insert a paragraph with a valid LINQ Reporting tag (outside the chart).
        builder.Writeln("<<[model.Title]>>");

        // Insert a table that will be populated by a foreach loop (also outside the chart).
        builder.Writeln("<<foreach [item in model.Items]>>");
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("<<[item.Category]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Value]>>");
        builder.EndRow();
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        doc.Save(TemplatePath);
    }

    private static void ValidateNoProhibitedTagsInCharts(Document doc)
    {
        // Tags that are not allowed inside chart elements.
        string[] prohibitedPrefixes = { "<<image", "<<bookmark", "<<link" };

        // Find all shapes in the document.
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapes)
        {
            if (shape.HasChart)
            {
                // Get any inner text of the shape (charts normally have none).
                string shapeText = shape.GetText();

                // Scan the shape's text for prohibited tags.
                foreach (string prefix in prohibitedPrefixes)
                {
                    if (shapeText.Contains(prefix, StringComparison.Ordinal))
                    {
                        throw new InvalidOperationException(
                            $"Prohibited tag '{prefix}' found inside a chart shape.");
                    }
                }
            }
        }
    }
}
