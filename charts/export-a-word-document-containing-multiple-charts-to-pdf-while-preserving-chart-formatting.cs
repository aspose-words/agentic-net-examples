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
        // Create a new blank Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -------------------------------------------------
        // Insert the first chart – a Column chart.
        // -------------------------------------------------
        Shape columnShape = builder.InsertChart(ChartType.Column, 400, 300);
        if (columnShape.HasChart)
        {
            Chart columnChart = columnShape.Chart;

            // Set chart title.
            columnChart.Title.Text = "Sales by Region";
            columnChart.Title.Show = true;
            columnChart.Title.Font.Size = 14;
            columnChart.Title.Font.Color = Color.DarkBlue;

            // Remove the demo data series.
            columnChart.Series.Clear();

            // Define categories and two series of values.
            string[] categories = { "North", "South", "East", "West" };
            columnChart.Series.Add("2019", categories, new double[] { 120, 150, 100, 130 });
            columnChart.Series.Add("2020", categories, new double[] { 140, 160, 110, 150 });

            // Enable data labels for each series and format them.
            foreach (ChartSeries series in columnChart.Series)
            {
                series.HasDataLabels = true;
                series.DataLabels.ShowValue = true;
                series.DataLabels.NumberFormat.FormatCode = "#,##0";
            }

            // Position the legend on the right side.
            columnChart.Legend.Position = LegendPosition.Right;
        }

        // Add a blank paragraph to separate the charts.
        builder.Writeln();

        // -------------------------------------------------
        // Insert the second chart – a Pie chart.
        // -------------------------------------------------
        Shape pieShape = builder.InsertChart(ChartType.Pie, 400, 300);
        if (pieShape.HasChart)
        {
            Chart pieChart = pieShape.Chart;

            // Set chart title.
            pieChart.Title.Text = "Market Share";
            pieChart.Title.Show = true;
            pieChart.Title.Font.Size = 14;
            pieChart.Title.Font.Color = Color.DarkGreen;

            // Remove the demo data series.
            pieChart.Series.Clear();

            // Define categories and a single series of values.
            string[] categoriesPie = { "Product A", "Product B", "Product C" };
            pieChart.Series.Add("Share", categoriesPie, new double[] { 45, 30, 25 });

            // Configure data labels for the pie series.
            ChartSeries pieSeries = pieChart.Series[0];
            pieSeries.HasDataLabels = true;
            pieSeries.DataLabels.ShowCategoryName = true;
            pieSeries.DataLabels.ShowValue = true;
            pieSeries.DataLabels.ShowPercentage = true;
            pieSeries.DataLabels.NumberFormat.FormatCode = "0.0%";

            // Position the legend at the bottom.
            pieChart.Legend.Position = LegendPosition.Bottom;
        }

        // -------------------------------------------------
        // Save the document as DOCX and then export to PDF.
        // -------------------------------------------------
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        string docxPath = Path.Combine(outputDir, "MultipleCharts.docx");
        doc.Save(docxPath); // Save the Word document with charts.

        string pdfPath = Path.Combine(outputDir, "MultipleCharts.pdf");
        doc.Save(pdfPath, SaveFormat.Pdf); // Export to PDF preserving chart formatting and data labels.
    }
}
