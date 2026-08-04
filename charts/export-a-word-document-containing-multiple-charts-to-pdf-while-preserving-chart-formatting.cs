using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -------------------------------------------------
        // First chart: Column chart with two series.
        // -------------------------------------------------
        Shape columnChartShape = builder.InsertChart(ChartType.Column, 400, 300);
        Chart columnChart = columnChartShape.Chart;

        // Set chart title.
        columnChart.Title.Text = "Sales by Quarter";
        columnChart.Title.Show = true;

        // Clear the demo data.
        columnChart.Series.Clear();

        // Define categories (X‑axis) and values for two series.
        string[] quarters = { "Q1", "Q2", "Q3", "Q4" };
        columnChart.Series.Add("2019", quarters, new double[] { 15000, 20000, 18000, 22000 });
        columnChart.Series.Add("2020", quarters, new double[] { 17000, 21000, 19000, 23000 });

        // Enable data labels for each series and configure them.
        foreach (ChartSeries series in columnChart.Series)
        {
            series.HasDataLabels = true;
            for (int i = 0; i < series.DataLabels.Count; i++)
            {
                series.DataLabels[i].ShowValue = true;
                series.DataLabels[i].ShowCategoryName = true;
                series.DataLabels[i].NumberFormat.FormatCode = "#,##0";
            }
        }

        // -------------------------------------------------
        // Second chart: Pie chart with a single series.
        // -------------------------------------------------
        builder.Writeln(); // Add a paragraph break between charts.
        Shape pieChartShape = builder.InsertChart(ChartType.Pie, 400, 300);
        Chart pieChart = pieChartShape.Chart;

        // Set chart title.
        pieChart.Title.Text = "Market Share";
        pieChart.Title.Show = true;

        // Clear the demo data.
        pieChart.Series.Clear();

        // Define categories and values.
        string[] products = { "Product A", "Product B", "Product C" };
        pieChart.Series.Add("Share", products, new double[] { 45.0, 30.0, 25.0 });

        // Enable data labels for the series.
        ChartSeries pieSeries = pieChart.Series[0];
        pieSeries.HasDataLabels = true;
        for (int i = 0; i < pieSeries.DataLabels.Count; i++)
        {
            pieSeries.DataLabels[i].ShowValue = true;
            pieSeries.DataLabels[i].ShowCategoryName = true;
            pieSeries.DataLabels[i].NumberFormat.FormatCode = "0.0%";
        }

        // -------------------------------------------------
        // Save the document containing the charts.
        // -------------------------------------------------
        string docPath = Path.Combine(Directory.GetCurrentDirectory(), "MultipleCharts.docx");
        doc.Save(docPath);

        // -------------------------------------------------
        // Export the same document to PDF, preserving formatting and data labels.
        // -------------------------------------------------
        string pdfPath = Path.Combine(Directory.GetCurrentDirectory(), "MultipleCharts.pdf");
        doc.Save(pdfPath, SaveFormat.Pdf);
    }
}
