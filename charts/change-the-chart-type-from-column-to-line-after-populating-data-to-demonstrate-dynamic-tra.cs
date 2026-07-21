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

        // Insert a column chart.
        Shape columnChartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart columnChart = columnChartShape.Chart;

        // Clear any demo data.
        columnChart.Series.Clear();

        // Sample data.
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };
        double[] sales = { 1500, 2300, 1800, 2100 };
        double[] profit = { 300, 500, 400, 450 };

        // Populate the column chart.
        columnChart.Series.Add("Sales", categories, sales);
        columnChart.Series.Add("Profit", categories, profit);

        // ----- Dynamic transformation: replace column chart with a line chart -----
        // Remove the original column chart shape.
        columnChartShape.Remove();

        // Insert a line chart at the same location.
        Shape lineChartShape = builder.InsertChart(ChartType.Line, 500, 300);
        Chart lineChart = lineChartShape.Chart;

        // Populate the line chart with the same data.
        lineChart.Series.Clear();
        lineChart.Series.Add("Sales", categories, sales);
        lineChart.Series.Add("Profit", categories, profit);

        // Save the document.
        doc.Save("DynamicChart.docx");
    }
}
