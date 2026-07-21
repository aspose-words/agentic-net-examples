using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
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
        // Insert the first chart: a column chart.
        // -------------------------------------------------
        Shape columnChartShape = builder.InsertChart(ChartType.Column, 432, 252);
        if (columnChartShape.HasChart)
        {
            Chart columnChart = columnChartShape.Chart;

            // Remove the demo data series.
            columnChart.Series.Clear();

            // Define categories and series data.
            string[] categories = { "Q1", "Q2", "Q3", "Q4" };
            columnChart.Series.Add("Product A", categories, new double[] { 120, 150, 170, 130 });
            columnChart.Series.Add("Product B", categories, new double[] { 80, 110, 130, 90 });

            // Set chart title.
            columnChart.Title.Text = "Quarterly Sales";
            columnChart.Title.Show = true;

            // Enable data labels for each series and format them.
            foreach (ChartSeries series in columnChart.Series)
            {
                series.HasDataLabels = true;
                for (int i = 0; i < series.DataLabels.Count; i++)
                {
                    series.DataLabels[i].ShowValue = true;
                    series.DataLabels[i].NumberFormat.FormatCode = "#,##0";
                }
            }

            // Apply custom colors to the series.
            columnChart.Series[0].Format.Fill.ForeColor = Color.CornflowerBlue;
            columnChart.Series[1].Format.Fill.ForeColor = Color.Orange;
        }

        // Add a paragraph break between the charts.
        builder.Writeln();

        // -------------------------------------------------
        // Insert the second chart: a pie chart.
        // -------------------------------------------------
        Shape pieChartShape = builder.InsertChart(ChartType.Pie, 432, 252);
        if (pieChartShape.HasChart)
        {
            Chart pieChart = pieChartShape.Chart;

            // Remove the demo data series.
            pieChart.Series.Clear();

            // Define categories and values.
            string[] categories = { "Apples", "Bananas", "Cherries", "Dates" };
            pieChart.Series.Add("Fruits", categories, new double[] { 30, 20, 25, 25 });

            // Set chart title.
            pieChart.Title.Text = "Fruit Distribution";
            pieChart.Title.Show = true;

            // Enable data labels and configure them.
            ChartSeries pieSeries = pieChart.Series[0];
            pieSeries.HasDataLabels = true;
            for (int i = 0; i < pieSeries.DataLabels.Count; i++)
            {
                pieSeries.DataLabels[i].ShowCategoryName = true;
                pieSeries.DataLabels[i].ShowValue = true;
                pieSeries.DataLabels[i].NumberFormat.FormatCode = "#,##0";
            }

            // Apply a base fill color for the series (individual slice colors can be set per point if needed).
            pieSeries.Format.Fill.ForeColor = Color.LightSalmon;
        }

        // -------------------------------------------------
        // Save the document in DOCX format.
        // -------------------------------------------------
        string docxPath = "MultipleCharts.docx";
        doc.Save(docxPath);

        // -------------------------------------------------
        // Export the same document to PDF while preserving chart formatting.
        // -------------------------------------------------
        string pdfPath = "MultipleCharts.pdf";
        doc.Save(pdfPath, SaveFormat.Pdf);
    }
}
