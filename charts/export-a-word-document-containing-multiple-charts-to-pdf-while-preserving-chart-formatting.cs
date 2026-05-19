using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using Aspose.Words.Saving;

public class ExportChartsToPdf
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ---------- First chart: Column chart ----------
        // Insert the chart with the desired style using the overload that accepts ChartStyle.
        Shape columnChartShape = builder.InsertChart(ChartType.Column, 432, 252, ChartStyle.Saturated);
        Chart columnChart = columnChartShape.Chart;

        // Set chart title.
        columnChart.Title.Text = "Sales by Quarter";
        columnChart.Title.Show = true;

        // Remove the demo data.
        columnChart.Series.Clear();

        // Define categories (X‑axis) and two series of values.
        string[] quarters = { "Q1", "Q2", "Q3", "Q4" };
        columnChart.Series.Add("2019", quarters, new double[] { 150, 200, 180, 220 });
        columnChart.Series.Add("2020", quarters, new double[] { 170, 210, 190, 230 });

        // Enable data labels for each series and configure them.
        foreach (ChartSeries series in columnChart.Series)
        {
            series.HasDataLabels = true;
            for (int i = 0; i < series.DataLabels.Count; i++)
            {
                ChartDataLabel label = series.DataLabels[i];
                label.ShowValue = true;
                label.NumberFormat.FormatCode = "0";
                label.Font.Size = 9;
                label.Font.Color = Color.Black;
            }
        }

        // Set legend position.
        columnChart.Legend.Position = LegendPosition.Bottom;

        // ---------- Second chart: Pie chart ----------
        // Move the builder to a new paragraph before inserting the next chart.
        builder.Writeln();

        // Insert the pie chart with the desired style.
        Shape pieChartShape = builder.InsertChart(ChartType.Pie, 432, 252, ChartStyle.Shaded);
        Chart pieChart = pieChartShape.Chart;

        // Set chart title.
        pieChart.Title.Text = "Market Share";
        pieChart.Title.Show = true;

        // Remove the demo data.
        pieChart.Series.Clear();

        // Add a single series with categories and values.
        string[] products = { "Product A", "Product B", "Product C" };
        pieChart.Series.Add("Share", products, new double[] { 45, 30, 25 });

        // Configure data labels for the pie series.
        ChartSeries pieSeries = pieChart.Series[0];
        pieSeries.HasDataLabels = true;
        for (int i = 0; i < pieSeries.DataLabels.Count; i++)
        {
            ChartDataLabel label = pieSeries.DataLabels[i];
            label.ShowCategoryName = true;
            label.ShowValue = true;
            label.NumberFormat.FormatCode = "0%";
            label.Font.Size = 9;
            label.Font.Color = Color.Black;
        }

        // Set legend position.
        pieChart.Legend.Position = LegendPosition.Right;

        // ---------- Save the document ----------
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        string docxPath = Path.Combine(outputDir, "Charts.docx");
        doc.Save(docxPath);

        // ---------- Export to PDF ----------
        string pdfPath = Path.Combine(outputDir, "Charts.pdf");
        doc.Save(pdfPath, SaveFormat.Pdf);
    }
}
