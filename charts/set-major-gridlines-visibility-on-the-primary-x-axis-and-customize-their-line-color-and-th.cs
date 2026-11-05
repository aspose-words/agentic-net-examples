using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart into the document.
        Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart chart = chartShape.Chart;

        // Access the primary X‑axis.
        ChartAxis xAxis = chart.AxisX;

        // Use reflection / dynamic to access gridlines in a version‑agnostic way.
        var majorGridlinesProp = xAxis.GetType().GetProperty("MajorGridlines");
        if (majorGridlinesProp != null)
        {
            dynamic majorGridlines = majorGridlinesProp.GetValue(xAxis);
            if (majorGridlines != null)
            {
                // Make the major gridlines visible.
                majorGridlines.Visible = true;

                // Customize line color and thickness.
                majorGridlines.LineFormat.Color = Color.Blue;
                majorGridlines.LineFormat.Width = 2.0; // Width in points.
            }
        }

        // Save the document with the modified chart.
        doc.Save("ChartWithCustomGridlines.docx");
    }
}
