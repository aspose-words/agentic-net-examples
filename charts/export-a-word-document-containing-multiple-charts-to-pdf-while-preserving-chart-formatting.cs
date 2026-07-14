using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;               // Needed for Shape
using Aspose.Words.Drawing.Charts;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -------------------------------------------------
        // First chart: Column chart with data labels.
        // -------------------------------------------------
        Shape columnShape = builder.InsertChart(ChartType.Column, 400, 300);
        if (columnShape.HasChart)
        {
            Chart columnChart = columnShape.Chart;

            // Remove the demo data.
            columnChart.Series.Clear();

            // Define categories and values.
            string[] categories = { "Q1", "Q2", "Q3", "Q4" };
            double[] values = { 15000, 20000, 18000, 22000 };

            // Add a series.
            columnChart.Series.Add("Revenue", categories, values);

            // Enable data labels and show the value for each point.
            ChartSeries series = columnChart.Series[0];
            series.HasDataLabels = true;
            for (int i = 0; i < series.DataLabels.Count; i++)
            {
                series.DataLabels[i].ShowValue = true;
            }

            // Set chart title.
            columnChart.Title.Text = "Quarterly Revenue";
            columnChart.Title.Show = true;

            // Position the legend.
            columnChart.Legend.Position = LegendPosition.Right;

            // Apply a simple fill color to the series.
            series.Format.Fill.ForeColor = Color.Blue;
        }

        // Add a paragraph break between charts.
        builder.Writeln();

        // -------------------------------------------------
        // Second chart: Pie chart with percentage labels.
        // -------------------------------------------------
        Shape pieShape = builder.InsertChart(ChartType.Pie, 300, 300);
        if (pieShape.HasChart)
        {
            Chart pieChart = pieShape.Chart;
            pieChart.Series.Clear();

            string[] productNames = { "Product A", "Product B", "Product C" };
            double[] marketShare = { 45, 30, 25 };

            pieChart.Series.Add("Market Share", productNames, marketShare);

            ChartSeries pieSeries = pieChart.Series[0];
            pieSeries.HasDataLabels = true;
            for (int i = 0; i < pieSeries.DataLabels.Count; i++)
            {
                pieSeries.DataLabels[i].ShowPercentage = true;
                pieSeries.DataLabels[i].ShowCategoryName = true;
            }

            pieChart.Title.Text = "Product Market Share";
            pieChart.Title.Show = true;

            pieChart.Legend.Position = LegendPosition.Bottom;
        }

        // Add another paragraph break.
        builder.Writeln();

        // -------------------------------------------------
        // Third chart: Line chart with category and value labels.
        // -------------------------------------------------
        Shape lineShape = builder.InsertChart(ChartType.Line, 400, 300);
        if (lineShape.HasChart)
        {
            Chart lineChart = lineShape.Chart;
            lineChart.Series.Clear();

            string[] months = { "Jan", "Feb", "Mar", "Apr", "May" };
            double[] temperatures = { 30, 35, 40, 45, 50 };

            lineChart.Series.Add("Temperature", months, temperatures);

            ChartSeries lineSeries = lineChart.Series[0];
            lineSeries.HasDataLabels = true;
            for (int i = 0; i < lineSeries.DataLabels.Count; i++)
            {
                lineSeries.DataLabels[i].ShowCategoryName = true;
                lineSeries.DataLabels[i].ShowValue = true;
            }

            lineChart.Title.Text = "Monthly Temperatures";
            lineChart.Title.Show = true;

            lineChart.Legend.Position = LegendPosition.Top;
        }

        // Save the document with charts.
        doc.Save("MultipleCharts.docx");

        // Export the same document to PDF, preserving chart formatting.
        doc.Save("MultipleCharts.pdf", SaveFormat.Pdf);
    }
}
