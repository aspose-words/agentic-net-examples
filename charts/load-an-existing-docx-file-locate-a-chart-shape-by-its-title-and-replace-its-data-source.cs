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
        // Step 1: Create a sample DOCX with a chart and a title.
        Document inputDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(inputDoc);

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        Chart chart = chartShape.Chart;

        // Set chart title.
        chart.Title.Text = "Sales Overview";

        // Add initial series data.
        chart.Series.Clear();
        chart.Series.Add("Q1", new double[] { 15, 25, 35 });
        chart.Series.Add("Q2", new double[] { 20, 30, 40 });

        // Save the input document.
        string inputPath = "input.docx";
        inputDoc.Save(inputPath);

        // Step 2: Load the existing DOCX.
        Document doc = new Document(inputPath);

        // Step 3: Locate the chart shape by its title.
        Shape? targetChartShape = doc.GetChildNodes(NodeType.Shape, true)
            .OfType<Shape>()
            .FirstOrDefault(s => s.HasChart && s.Chart.Title != null && s.Chart.Title.Text == "Sales Overview");

        if (targetChartShape == null)
        {
            throw new InvalidOperationException("Chart with the specified title was not found.");
        }

        // Step 4: Replace the chart's data source.
        Chart targetChart = targetChartShape.Chart;
        targetChart.Series.Clear();

        // New data series.
        targetChart.Series.Add("2023", new double[] { 50, 60, 70 });
        targetChart.Series.Add("2024", new double[] { 55, 65, 75 });

        // Optionally update the title to reflect new data.
        targetChart.Title.Text = "Projected Sales";

        // Step 5: Save the updated document.
        string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}
