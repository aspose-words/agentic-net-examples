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
        // Step 1: Create a sample DOCX with a chart that has a known title.
        const string inputPath = "input.docx";
        CreateSampleDocument(inputPath);

        // Step 2: Load the existing document.
        Document doc = new Document(inputPath);

        // Step 3: Locate the chart shape by its title.
        const string targetTitle = "Sample Chart";
        Shape? chartShape = doc.GetChildNodes(NodeType.Shape, true)
                               .OfType<Shape>()
                               .FirstOrDefault(s => s.HasChart && s.Chart.Title.Text == targetTitle);

        if (chartShape == null)
        {
            throw new InvalidOperationException($"No chart with title '{targetTitle}' was found.");
        }

        // Step 4: Replace the chart's data source.
        Chart chart = chartShape.Chart;
        chart.Series.Clear(); // Remove existing demo data.

        // New data for the chart.
        string[] categories = { "Category A", "Category B", "Category C" };
        double[] values = { 12.5, 23.0, 34.5 };

        // Add a new series with the new data.
        chart.Series.Add("Updated Series", categories, values);

        // Optional: Update the chart title to reflect the change.
        chart.Title.Text = "Updated Chart";

        // Step 5: Save the updated document.
        const string outputPath = "updated.docx";
        doc.Save(outputPath);
    }

    // Helper method to create a DOCX containing a chart with a specific title.
    private static void CreateSampleDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart.
        Shape shape = builder.InsertChart(ChartType.Column, 432, 252);
        Chart chart = shape.Chart;

        // Set the chart title.
        chart.Title.Text = "Sample Chart";
        chart.Title.Show = true;

        // Add some initial demo data (optional, will be replaced later).
        string[] demoCategories = { "Q1", "Q2", "Q3", "Q4" };
        double[] demoValues = { 10, 20, 30, 40 };
        chart.Series.Clear(); // Ensure a clean start.
        chart.Series.Add("Demo Series", demoCategories, demoValues);

        // Save the document.
        doc.Save(filePath);
    }
}
