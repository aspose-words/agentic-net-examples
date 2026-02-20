using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a Bar chart shape with a width of 400 points and a height of 300 points.
        Shape chartShape = builder.InsertChart(ChartType.Bar, 400, 300);
        Chart chart = chartShape.Chart;

        // Set the chart title and make it visible.
        ChartTitle title = chart.Title;
        title.Text = "Sales Overview";
        title.Font.Size = 14;
        title.Font.Color = Color.Blue;
        title.Show = true;

        // Define categories for the X axis.
        string[] categories = new string[] { "Q1", "Q2", "Q3", "Q4" };

        // Clear any default series and add custom data series.
        chart.Series.Clear();
        chart.Series.Add("2019", categories, new double[] { 120, 150, 170, 200 });
        chart.Series.Add("2020", categories, new double[] { 130, 160, 180, 210 });

        // Save the document to a DOCX file.
        doc.Save("ChartInShape.docx");
    }
}
