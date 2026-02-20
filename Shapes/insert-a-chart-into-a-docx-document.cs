using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

class InsertChartExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart with the specified width and height.
        // The InsertChart method returns a Shape that contains the chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 400, 300);

        // Access the Chart object from the shape.
        Chart chart = chartShape.Chart;

        // Optional: set a title for the chart.
        ChartTitle title = chart.Title;
        title.Text = "Sales Overview";
        title.Font.Size = 14;
        title.Font.Color = Color.DarkBlue;
        title.Show = true;      // Ensure the title is visible.
        title.Overlay = false; // Do not allow other elements to overlap the title.

        // Optional: clear the default demo series and add custom data.
        chart.Series.Clear();

        string[] categories = { "Q1", "Q2", "Q3", "Q4" };
        chart.Series.Add("Product A", categories, new double[] { 12000, 15000, 13000, 17000 });
        chart.Series.Add("Product B", categories, new double[] { 10000, 14000, 11000, 16000 });

        // Save the document to a DOCX file.
        doc.Save("ChartDocument.docx");
    }
}
