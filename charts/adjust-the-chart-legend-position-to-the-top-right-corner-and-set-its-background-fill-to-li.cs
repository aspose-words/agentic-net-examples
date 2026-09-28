using System;
using System.Drawing;
using System.Linq;
using System.Reflection;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        if (!chartShape.HasChart)
            throw new InvalidOperationException("The inserted shape does not contain a chart.");

        Chart chart = chartShape.Chart;

        // Ensure the chart has at least one series.
        chart.Series.Clear();
        chart.Series.Add("Series 1", new double[] { 10, 20, 30 });

        // Adjust the legend: position to top right.
        ChartLegend legend = chart.Legend;
        legend.Position = LegendPosition.TopRight;

        // Set legend background fill to light gray, using reflection to stay compatible with
        // Aspose.Words versions that may not expose FillFormat directly.
        PropertyInfo fillProp = typeof(ChartLegend).GetProperty("FillFormat");
        if (fillProp != null)
        {
            object fillObj = fillProp.GetValue(legend);
            if (fillObj != null)
            {
                PropertyInfo foreColorProp = fillObj.GetType().GetProperty("ForeColor");
                if (foreColorProp != null && foreColorProp.CanWrite)
                {
                    foreColorProp.SetValue(fillObj, Color.LightGray);
                }
            }
        }

        // Save the document.
        doc.Save("ChartLegendModified.docx");
    }
}
