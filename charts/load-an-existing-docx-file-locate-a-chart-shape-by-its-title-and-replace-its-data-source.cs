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
        // Paths for the sample input and output documents.
        string inputPath = "input.docx";
        string outputPath = "output.docx";

        // -----------------------------------------------------------------
        // Step 1: Create a sample DOCX with a chart if it does not already exist.
        // -----------------------------------------------------------------
        if (!File.Exists(inputPath))
        {
            Document createDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(createDoc);

            // Insert a column chart.
            Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
            // Use the shape's Title property to identify the chart later.
            chartShape.Title = "TargetChart";

            // Optionally set the chart's internal title.
            Chart chart = chartShape.Chart;
            chart.Title.Text = "Original Chart Title";

            // Populate the chart with initial data.
            chart.Series.Clear();
            chart.Series.Add("InitialSeries", new double[] { 5, 10, 15 });

            // Save the document that will be used as input.
            createDoc.Save(inputPath);
        }

        // -----------------------------------------------------------------
        // Step 2: Load the existing document.
        // -----------------------------------------------------------------
        Document doc = new Document(inputPath);

        // -----------------------------------------------------------------
        // Step 3: Locate the chart shape by its title.
        // -----------------------------------------------------------------
        Shape targetChartShape = doc.GetChildNodes(NodeType.Shape, true)
            .OfType<Shape>()
            .FirstOrDefault(s => s.HasChart && s.Title == "TargetChart");

        if (targetChartShape == null)
            throw new InvalidOperationException("Chart with the specified title was not found.");

        // -----------------------------------------------------------------
        // Step 4: Replace the chart's data source.
        // -----------------------------------------------------------------
        Chart targetChart = targetChartShape.Chart;

        // Clear existing series.
        targetChart.Series.Clear();

        // Add new series with new data.
        targetChart.Series.Add("NewSeries", new double[] { 12, 24, 36, 48 });

        // Optionally update the chart's internal title to reflect the change.
        targetChart.Title.Text = "Updated Chart Title";

        // -----------------------------------------------------------------
        // Step 5: Save the modified document.
        // -----------------------------------------------------------------
        doc.Save(outputPath);
    }
}
