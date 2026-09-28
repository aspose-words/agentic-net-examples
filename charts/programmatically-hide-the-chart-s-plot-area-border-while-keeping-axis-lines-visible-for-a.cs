using System;
using System.Reflection;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class HideChartPlotAreaBorder
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);

        // Ensure the shape actually contains a chart.
        if (!chartShape.HasChart)
            throw new InvalidOperationException("The inserted shape does not contain a chart.");

        // Access the chart.
        Chart chart = chartShape.Chart;

        // Hide the plot area border while keeping axis lines visible.
        // Use reflection to stay compatible with versions where PlotArea is not exposed directly.
        PropertyInfo? plotAreaProp = chart.GetType().GetProperty("PlotArea", BindingFlags.Public | BindingFlags.Instance);
        if (plotAreaProp != null)
        {
            object? plotArea = plotAreaProp.GetValue(chart);
            if (plotArea != null)
            {
                PropertyInfo? borderProp = plotArea.GetType().GetProperty("Border", BindingFlags.Public | BindingFlags.Instance);
                object? border = borderProp?.GetValue(plotArea);
                if (border != null)
                {
                    PropertyInfo? visibleProp = border.GetType().GetProperty("Visible", BindingFlags.Public | BindingFlags.Instance);
                    if (visibleProp != null && visibleProp.CanWrite)
                    {
                        visibleProp.SetValue(border, false);
                    }
                }
            }
        }

        // Save the document.
        string outputPath = "HidePlotAreaBorder.docx";
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to {outputPath}");
    }
}
