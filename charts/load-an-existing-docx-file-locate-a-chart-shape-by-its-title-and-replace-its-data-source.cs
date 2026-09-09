using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Define file names in the working directory.
        const string inputPath = "input.docx";
        const string outputPath = "updated.docx";

        // -----------------------------------------------------------------
        // Step 1: Create a sample DOCX with a chart if it does not exist.
        // -----------------------------------------------------------------
        if (!File.Exists(inputPath))
        {
            // Create a blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a column chart.
            Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
            Chart chart = chartShape.Chart;

            // Give the chart a title that we will later search for.
            chart.Title.Text = "Sales Chart";
            chart.Title.Show = true;

            // Save the document that will serve as the input.
            doc.Save(inputPath);
        }

        // ---------------------------------------------------------------
        // Step 2: Load the existing document.
        // ---------------------------------------------------------------
        Document loadedDoc = new Document(inputPath);

        // ---------------------------------------------------------------
        // Step 3: Locate the chart shape by its title.
        // ---------------------------------------------------------------
        Shape? targetShape = null;
        foreach (Shape shape in loadedDoc.GetChildNodes(NodeType.Shape, true))
        {
            if (!shape.HasChart) continue;

            Chart chart = shape.Chart;
            // Ensure the title object exists before accessing its Text.
            if (chart.Title != null && chart.Title.Text == "Sales Chart")
            {
                targetShape = shape;
                break;
            }
        }

        if (targetShape == null)
        {
            throw new InvalidOperationException("Chart with the specified title was not found.");
        }

        // ---------------------------------------------------------------
        // Step 4: Replace the chart's data source.
        // ---------------------------------------------------------------
        Chart targetChart = targetShape.Chart;

        // Clear any existing series.
        targetChart.Series.Clear();

        // Define new categories and values.
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };
        double[] values = { 15.0, 25.0, 35.0, 45.0 };

        // Add a new series with the new data.
        targetChart.Series.Add("New Series", categories, values);

        // Optionally update the chart title to reflect the change.
        targetChart.Title.Text = "Updated Sales Chart";

        // ---------------------------------------------------------------
        // Step 5: Save the modified document.
        // ---------------------------------------------------------------
        loadedDoc.Save(outputPath);
    }
}
