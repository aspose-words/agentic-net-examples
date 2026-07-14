using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart.
        Shape columnShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart columnChart = columnShape.Chart;

        // Clear any demo data.
        columnChart.Series.Clear();

        // Define categories and values.
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };
        double[] sales = { 120, 150, 170, 130 };
        double[] profit = { 30, 45, 50, 35 };

        // Populate the column chart.
        columnChart.Series.Add("Sales", categories, sales);
        columnChart.Series.Add("Profit", categories, profit);
        columnChart.Title.Text = "Sales and Profit (Column Chart)";
        columnChart.Title.Show = true;

        // Insert a line chart at the same position to demonstrate dynamic transformation.
        builder.MoveTo(columnShape);
        Shape lineShape = builder.InsertChart(ChartType.Line, 500, 300);
        Chart lineChart = lineShape.Chart;

        // Populate the line chart with the same data.
        lineChart.Series.Clear();
        lineChart.Series.Add("Sales", categories, sales);
        lineChart.Series.Add("Profit", categories, profit);
        lineChart.Title.Text = "Sales and Profit (Line Chart)";
        lineChart.Title.Show = true;

        // Remove the original column chart.
        columnShape.Remove();

        // Save the document.
        doc.Save("DynamicChartTransformation.docx");
    }
}
