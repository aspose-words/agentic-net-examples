using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the demo data.
        chart.Series.Clear();

        // Define categories (X‑axis) and values (Y‑axis).
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };
        double[] values = { 120.5, 150.0, 130.75, 170.25 };

        // Add a single series with the data.
        chart.Series.Add("Sales", categories, values);

        // ----- Dynamic transformation: replace the column chart with a line chart -----
        // Move the builder to the position of the existing chart shape.
        builder.MoveTo(chartShape);
        // Remove the old chart shape.
        chartShape.Remove();

        // Insert a new line chart at the same location.
        Shape lineChartShape = builder.InsertChart(ChartType.Line, 500, 300);
        Chart lineChart = lineChartShape.Chart;

        // Populate the line chart with the same data.
        lineChart.Series.Clear();
        lineChart.Series.Add("Sales", categories, values);

        // Save the document.
        doc.Save("DynamicChartTransformation.docx");
    }
}
