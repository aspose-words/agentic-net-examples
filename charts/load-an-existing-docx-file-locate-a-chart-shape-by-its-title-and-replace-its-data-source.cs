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
        // Step 1: Create a sample DOCX with a chart that has a title.
        Document inputDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(inputDoc);

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        Chart chart = chartShape.Chart;

        // Set the chart title which will be used to locate the chart later.
        chart.Title.Text = "Sales Chart";

        // Add an initial series (optional, just to have some data).
        chart.Series.Clear();
        chart.Series.Add("Initial Series", new double[] { 10, 20, 30 });

        // Save the input document.
        string inputPath = "input.docx";
        inputDoc.Save(inputPath);

        // Step 2: Load the existing DOCX.
        Document doc = new Document(inputPath);

        // Step 3: Locate the chart shape by its title.
        Shape targetShape = doc.GetChildNodes(NodeType.Shape, true)
            .OfType<Shape>()
            .FirstOrDefault(s => s.HasChart && s.Chart.Title != null && s.Chart.Title.Text == "Sales Chart");

        if (targetShape == null)
        {
            throw new InvalidOperationException("Chart with the specified title was not found.");
        }

        // Step 4: Replace the chart's data source.
        Chart targetChart = targetShape.Chart;
        targetChart.Series.Clear();

        // Add new series with new data.
        targetChart.Series.Add("Q1", new double[] { 15, 25, 35 });
        targetChart.Series.Add("Q2", new double[] { 20, 30, 40 });

        // Optionally, update the chart title to reflect the change.
        targetChart.Title.Text = "Updated Sales Chart";

        // Step 5: Save the updated document.
        string outputPath = "updated.docx";
        doc.Save(outputPath);
    }
}
