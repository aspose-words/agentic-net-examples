using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX with a chart that has a known title.
        Document inputDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(inputDoc);

        // Insert a column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        Chart chart = chartShape.Chart;

        // Set the chart title.
        chart.Title.Text = "Sales Chart";
        chart.Title.Show = true;

        // Populate the chart with initial data.
        chart.Series.Clear();
        string[] initialCategories = { "Q1", "Q2", "Q3", "Q4" };
        chart.Series.Add("2022", initialCategories, new double[] { 10, 20, 30, 40 });

        // Save the input document.
        const string inputPath = "input.docx";
        inputDoc.Save(inputPath);

        // Load the existing document.
        Document doc = new Document(inputPath);

        // Locate the chart shape by its title.
        Shape targetShape = null;
        foreach (Node node in doc.GetChildNodes(NodeType.Shape, true))
        {
            if (node is Shape shape && shape.HasChart && shape.Chart.Title.Text == "Sales Chart")
            {
                targetShape = shape;
                break;
            }
        }

        if (targetShape == null)
            throw new InvalidOperationException("Chart with the specified title was not found.");

        // Replace the chart's data source.
        Chart targetChart = targetShape.Chart;
        targetChart.Series.Clear();

        string[] newCategories = { "Jan", "Feb", "Mar", "Apr" };
        targetChart.Series.Add("2023", newCategories, new double[] { 15, 25, 35, 45 });

        // Save the updated document.
        const string outputPath = "updated.docx";
        doc.Save(outputPath);
    }
}
