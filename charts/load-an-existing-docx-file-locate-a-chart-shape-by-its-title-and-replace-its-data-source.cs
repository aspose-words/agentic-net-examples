using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Paths for the sample input and output documents.
        const string inputPath = "InputChart.docx";
        const string outputPath = "ModifiedChart.docx";

        // Ensure a sample document with a chart exists.
        if (!File.Exists(inputPath))
        {
            CreateSampleDocumentWithChart(inputPath);
        }

        // Load the existing document.
        Document doc = new Document(inputPath);

        // Locate the chart shape by its title.
        Shape? chartShape = FindChartShapeByTitle(doc, "Sales Chart");
        if (chartShape == null)
        {
            throw new InvalidOperationException("Chart with the specified title was not found.");
        }

        // Access the chart object.
        Chart chart = chartShape.Chart;

        // Replace the chart's data source with new data.
        ReplaceChartData(chart);

        // Save the modified document.
        doc.Save(outputPath);
    }

    // Creates a simple document containing a column chart with a title.
    private static void CreateSampleDocumentWithChart(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 400, 300);
        Chart chart = chartShape.Chart;

        // Set a title for later identification.
        ChartTitle title = chart.Title;
        title.Text = "Sales Chart";
        title.Show = true;

        // The chart already contains demo data; no further changes needed.
        doc.Save(filePath);
    }

    // Searches all shapes in the document for a chart whose title matches the given text.
    private static Shape? FindChartShapeByTitle(Document doc, string titleText)
    {
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapes)
        {
            if (shape.HasChart && shape.Chart.Title != null && shape.Chart.Title.Text == titleText)
            {
                return shape;
            }
        }
        return null;
    }

    // Clears existing series and adds a new series with custom categories and values.
    private static void ReplaceChartData(Chart chart)
    {
        // Remove all existing series.
        chart.Series.Clear();

        // Define new categories and values.
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };
        double[] values = { 150.0, 200.0, 180.0, 220.0 };

        // Add a new series to the chart.
        chart.Series.Add("Quarterly Sales", categories, values);
    }
}
