using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;          // <-- added
using Aspose.Words.Drawing.Charts;

class InsertAreaChartExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an Area chart with the desired size.
        // ChartType.Area corresponds to a standard 2‑D area chart.
        Shape chartShape = builder.InsertChart(ChartType.Area, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the demo series that Aspose.Words adds by default.
        chart.Series.Clear();

        // Define categories (X‑axis) – dates are typical for Area charts.
        DateTime[] dates = new DateTime[]
        {
            new DateTime(2021, 1, 1),
            new DateTime(2021, 2, 1),
            new DateTime(2021, 3, 1),
            new DateTime(2021, 4, 1),
            new DateTime(2021, 5, 1)
        };

        // Define Y‑axis values for a single series.
        double[] values = new double[] { 10.5, 14.2, 12.8, 18.3, 16.0 };

        // Add the series to the chart.
        chart.Series.Add("Sales", dates, values);

        // Optional: set a chart title.
        chart.Title.Text = "Quarterly Sales";
        chart.Title.Font.Size = 14;
        chart.Title.Font.Color = Color.DarkBlue;
        chart.Title.Show = true;

        // Optional: format the chart background.
        chart.Format.Fill.Solid(Color.LightYellow);

        // Save the document to a DOCX file.
        doc.Save("AreaChart.docx");
    }
}
