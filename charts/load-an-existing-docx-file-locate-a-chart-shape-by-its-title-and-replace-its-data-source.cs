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
        // Step 1: Create a sample DOCX with a chart and a specific title.
        Document inputDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(inputDoc);

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        Chart chart = chartShape.Chart;

        // Set the chart title which we will later use to locate the chart.
        chart.Title.Text = "Sales Chart";
        chart.Title.Show = true;

        // Add some initial demo data.
        chart.Series.Clear();
        chart.Series.Add("Initial Series",
            new[] { "Jan", "Feb", "Mar", "Apr" },
            new[] { 5.0, 7.5, 3.2, 6.8 });

        // Save the document that will be loaded later.
        const string inputPath = "input.docx";
        inputDoc.Save(inputPath);

        // Step 2: Load the existing document.
        Document doc = new Document(inputPath);

        // Step 3: Locate the chart shape by its title.
        Shape? targetShape = doc.GetChildNodes(NodeType.Shape, true)
            .OfType<Shape>()
            .FirstOrDefault(s => s.HasChart && s.Chart.Title.Text == "Sales Chart");

        if (targetShape == null)
            throw new InvalidOperationException("Chart with the specified title was not found.");

        // Step 4: Replace the chart's data source.
        Chart targetChart = targetShape.Chart;

        // Clear existing series.
        targetChart.Series.Clear();

        // Define new categories and values.
        string[] newCategories = { "Q1", "Q2", "Q3", "Q4" };
        double[] newValues = { 10.0, 20.0, 30.0, 40.0 };

        // Add a new series with the new data.
        targetChart.Series.Add("New Sales Data", newCategories, newValues);

        // Optionally, update the title to reflect the change.
        targetChart.Title.Text = "Updated Sales Chart";

        // Step 5: Save the updated document.
        const string outputPath = "updated.docx";
        doc.Save(outputPath);
    }
}
