using System;
using Aspose.Words;
using Aspose.Words.Drawing.Charts;
using Aspose.Words.Drawing;
using System.Drawing;

public class ChartCustomizer
{
    /// <summary>
    /// Updates the title text of the first chart in the document and toggles the legend visibility.
    /// </summary>
    /// <param name="inputPath">Path to the source .docx file containing a chart.</param>
    /// <param name="outputPath">Path where the modified document will be saved.</param>
    /// <param name="newTitle">The text to set as the chart title.</param>
    /// <param name="showLegend">If true, the legend will be shown; otherwise it will be hidden.</param>
    public static void UpdateChart(string inputPath, string outputPath, string newTitle, bool showLegend)
    {
        // Load the existing document.
        Document doc = new Document(inputPath);

        // Find the first chart shape in the document.
        Shape chartShape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
        if (chartShape == null || chartShape.Chart == null)
            throw new InvalidOperationException("No chart found in the document.");

        Chart chart = chartShape.Chart;

        // Update the chart title text and ensure the title is visible.
        ChartTitle title = chart.Title;
        title.Text = newTitle;
        title.Show = true; // Make sure the title is displayed.

        // Toggle legend visibility.
        ChartLegend legend = chart.Legend;
        if (showLegend)
        {
            // Show the legend at a default position (right side).
            legend.Position = LegendPosition.Right;
        }
        else
        {
            // Hide the legend.
            legend.Position = LegendPosition.None;
        }

        // Save the modified document.
        doc.Save(outputPath);
    }

    // Example usage.
    public static void Main()
    {
        string inputFile = @"C:\Docs\InputChart.docx";
        string outputFile = @"C:\Docs\OutputChart.docx";

        string desiredTitle = "Quarterly Sales Report";
        bool displayLegend = true; // Set to false to hide the legend.

        UpdateChart(inputFile, outputFile, desiredTitle, displayLegend);
    }
}
