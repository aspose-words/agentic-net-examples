using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX with a chart that has a title.
        const string inputPath = "input.docx";
        CreateSampleDocument(inputPath);

        // Load the existing document.
        Document doc = new Document(inputPath);

        // Locate the chart shape by its title.
        Shape chartShape = doc.GetChildNodes(NodeType.Shape, true)
                              .OfType<Shape>()
                              .FirstOrDefault(s => s.HasChart && s.Chart.Title.Text == "Sample Chart");

        if (chartShape == null)
            throw new InvalidOperationException("Chart with the specified title was not found.");

        // Replace the chart's data source.
        Chart chart = chartShape.Chart;
        chart.Series.Clear();

        string[] categories = { "Category A", "Category B", "Category C" };
        double[] values = { 15.0, 30.0, 45.0 };
        chart.Series.Add("Updated Series", categories, values);

        // Optionally update the chart title to reflect the change.
        chart.Title.Text = "Updated Chart";

        // Save the modified document.
        const string outputPath = "updated.docx";
        doc.Save(outputPath);
    }

    // Helper method to create a DOCX containing a chart with a title.
    private static void CreateSampleDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        Chart chart = chartShape.Chart;

        // Set a title for later identification.
        chart.Title.Text = "Sample Chart";
        chart.Title.Show = true;

        // Add initial demo data.
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };
        chart.Series.Clear();
        chart.Series.Add("Initial Series", categories, new double[] { 10, 20, 30, 40 });

        // Save the document.
        doc.Save(filePath);
    }
}
