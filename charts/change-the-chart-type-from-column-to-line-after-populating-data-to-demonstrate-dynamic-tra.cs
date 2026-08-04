using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a Column chart and obtain its Chart object.
        Shape columnChartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart columnChart = columnChartShape.Chart;

        // Remove the demo data that comes with a newly inserted chart.
        columnChart.Series.Clear();

        // Populate the chart with sample data.
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };
        double[] sales = { 15000, 21000, 18000, 24000 };
        double[] expenses = { 12000, 16000, 13000, 19000 };

        columnChart.Series.Add("Sales", categories, sales);
        columnChart.Series.Add("Expenses", categories, expenses);

        // ----- Dynamic transformation: replace the column chart with a line chart -----
        // Store the data that we want to keep.
        // (In this simple example we already have the arrays above.)

        // Insert a Line chart right after the column chart.
        // The builder is positioned after the previously inserted shape,
        // so the new chart will appear directly after it.
        Shape lineChartShape = builder.InsertChart(ChartType.Line, 500, 300);
        Chart lineChart = lineChartShape.Chart;

        // Clear any demo data in the new chart.
        lineChart.Series.Clear();

        // Add the same series data to the line chart.
        lineChart.Series.Add("Sales", categories, sales);
        lineChart.Series.Add("Expenses", categories, expenses);

        // Remove the original column chart from the document.
        columnChartShape.Remove();

        // Save the document.
        doc.Save("DynamicChartTransformation.docx");
    }
}
