using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class DynamicChartTransformation
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart.
        Shape columnChartShape = builder.InsertChart(ChartType.Column, 500, 300);
        Chart columnChart = columnChartShape.Chart;

        // Remove the demo data.
        columnChart.Series.Clear();

        // Populate the chart with sample data.
        string[] categories = { "Jan", "Feb", "Mar" };
        double[] values = { 100, 150, 130 };
        columnChart.Series.Add("Sales", categories, values);

        // ----- Dynamic transformation -----
        // Aspose.Words does not allow changing the ChartType of an existing chart.
        // To demonstrate a transformation, we create a new line chart and copy the data.

        // Remove the original column chart shape.
        columnChartShape.Remove();

        // Insert a line chart at the same position.
        Shape lineChartShape = builder.InsertChart(ChartType.Line, 500, 300);
        Chart lineChart = lineChartShape.Chart;

        // Clear any demo data that comes with the new chart.
        lineChart.Series.Clear();

        // Copy the previously created series into the line chart.
        lineChart.Series.Add("Sales", categories, values);

        // Optional: add a title to illustrate the transformation.
        lineChart.Title.Text = "Sales Over Time";
        lineChart.Title.Show = true;

        // Save the document.
        doc.Save("DynamicChartTransformation.docx");
    }
}
