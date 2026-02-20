using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;            // <-- added
using Aspose.Words.Drawing.Charts;

class InsertColumnChart
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart shape with the desired size.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Remove the automatically generated demo series.
        chart.Series.Clear();

        // Define categories for the X‑axis.
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };

        // Add two data series to the chart.
        chart.Series.Add("Revenue", categories, new double[] { 12000, 15000, 13000, 17000 });
        chart.Series.Add("Profit",  categories, new double[] { 3000, 4000, 3500, 5000 });

        // Set a visible title for the chart.
        chart.Title.Text = "Quarterly Financials";
        chart.Title.Show = true;
        chart.Title.Font.Size = 14;
        chart.Title.Font.Color = Color.DarkBlue;

        // Apply a predefined chart style (optional).
        chart.Style = ChartStyle.Shaded;

        // Save the document containing the column chart.
        doc.Save("ColumnChart.docx");
    }
}
