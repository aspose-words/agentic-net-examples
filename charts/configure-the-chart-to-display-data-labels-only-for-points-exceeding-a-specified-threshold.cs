using System;
using Aspose.Words;
using Aspose.Words.Drawing;               // Needed for Shape
using Aspose.Words.Drawing.Charts;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Prepare sample data.
        string[] categories = { "A", "B", "C", "D", "E" };
        double[] values = { 10, 25, 5, 30, 15 };
        double threshold = 20.0; // Labels will be shown only for values greater than this.

        // Remove the demo series and add our custom series.
        chart.Series.Clear();
        chart.Series.Add("Sample Series", categories, values);

        // Get the series we just added.
        ChartSeries series = chart.Series[0];

        // Enable data labels for the series.
        series.HasDataLabels = true;

        // Configure each data label based on the point's value.
        for (int i = 0; i < series.YValues.Count; i++)
        {
            double pointValue = series.YValues[i].DoubleValue;

            if (pointValue > threshold)
            {
                // Show the value for points exceeding the threshold.
                series.DataLabels[i].ShowValue = true;
                series.DataLabels[i].ShowCategoryName = false;
                series.DataLabels[i].ShowSeriesName = false;
                series.DataLabels[i].NumberFormat.FormatCode = "0.##";
                series.DataLabels[i].IsHidden = false;
            }
            else
            {
                // Hide the label for points below the threshold.
                series.DataLabels[i].ShowValue = false;
                series.DataLabels[i].IsHidden = true;
            }
        }

        // Save the document.
        doc.Save("ChartWithConditionalDataLabels.docx");
    }
}
