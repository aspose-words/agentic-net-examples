using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class ExportChartsToPdf
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -------------------- First Chart (Column) --------------------
        // Insert a column chart.
        Shape chartShape1 = builder.InsertChart(ChartType.Column, 400, 300);
        if (!chartShape1.HasChart)
            throw new InvalidOperationException("The inserted shape does not contain a chart.");

        Chart chart1 = chartShape1.Chart;

        // Remove the demo data.
        chart1.Series.Clear();

        // Define categories and values.
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };
        double[] sales2022 = { 15000, 21000, 18000, 24000 };
        double[] sales2023 = { 17000, 23000, 19000, 26000 };

        // Add two series.
        ChartSeries series2022 = chart1.Series.Add("2022", categories, sales2022);
        ChartSeries series2023 = chart1.Series.Add("2023", categories, sales2023);

        // Set series colors.
        series2022.Format.Fill.ForeColor = Color.CornflowerBlue;
        series2023.Format.Fill.ForeColor = Color.OrangeRed;

        // Enable and configure data labels for both series.
        foreach (ChartSeries series in chart1.Series)
        {
            series.HasDataLabels = true;
            for (int i = 0; i < series.DataLabels.Count; i++)
            {
                series.DataLabels[i].ShowValue = true;
                series.DataLabels[i].NumberFormat.FormatCode = "#,##0";
                series.DataLabels[i].Font.Size = 9;
                series.DataLabels[i].Font.Color = Color.Black;
            }
        }

        // Add a title.
        chart1.Title.Text = "Annual Sales Comparison";
        chart1.Title.Font.Size = 14;
        chart1.Title.Font.Color = Color.DarkBlue;
        chart1.Title.Show = true;

        // Position the legend.
        chart1.Legend.Position = LegendPosition.Bottom;
        chart1.Legend.Overlay = false;

        // -------------------- Second Chart (Pie) --------------------
        // Move to a new paragraph before inserting the next chart.
        builder.Writeln();
        // Insert a pie chart.
        Shape chartShape2 = builder.InsertChart(ChartType.Pie, 400, 300);
        if (!chartShape2.HasChart)
            throw new InvalidOperationException("The inserted shape does not contain a chart.");

        Chart chart2 = chartShape2.Chart;

        // Remove the demo data.
        chart2.Series.Clear();

        // Define categories and values for the pie chart.
        string[] productCategories = { "Software", "Hardware", "Services", "Support" };
        double[] revenue = { 45, 30, 15, 10 };

        // Add a single series.
        ChartSeries revenueSeries = chart2.Series.Add("Revenue Share", productCategories, revenue);

        // Set slice colors.
        revenueSeries.Format.Fill.ForeColor = Color.LightSeaGreen;

        // Enable data labels.
        revenueSeries.HasDataLabels = true;
        for (int i = 0; i < revenueSeries.DataLabels.Count; i++)
        {
            revenueSeries.DataLabels[i].ShowCategoryName = true;
            revenueSeries.DataLabels[i].ShowValue = true;
            revenueSeries.DataLabels[i].NumberFormat.FormatCode = "0%";
            revenueSeries.DataLabels[i].Separator = " - ";
            revenueSeries.DataLabels[i].Font.Size = 9;
            revenueSeries.DataLabels[i].Font.Color = Color.Black;
        }

        // Add a title.
        chart2.Title.Text = "Revenue Distribution";
        chart2.Title.Font.Size = 14;
        chart2.Title.Font.Color = Color.DarkGreen;
        chart2.Title.Show = true;

        // Position the legend.
        chart2.Legend.Position = LegendPosition.Right;
        chart2.Legend.Overlay = true;

        // -------------------- Save Document --------------------
        // Save the document with charts as DOCX.
        string docxPath = "MultipleCharts.docx";
        doc.Save(docxPath);

        // Export the same document to PDF while preserving chart formatting and data labels.
        string pdfPath = "MultipleCharts.pdf";
        doc.Save(pdfPath, SaveFormat.Pdf);
    }
}
