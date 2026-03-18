using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using Aspose.Words.Saving;

class Apply3DRotationToColumnChart
{
    static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a 3‑D column chart.
        Shape chartShape = builder.InsertChart(ChartType.Column3D, 500, 350);
        Chart chart = chartShape.Chart;

        // Remove the demo data that Aspose.Words inserts by default.
        chart.Series.Clear();

        // Add a simple series.
        chart.Series.Add("Sales", new[] { "Q1", "Q2", "Q3", "Q4" }, new[] { 120.0, 150.0, 180.0, 200.0 });

        // Set a title and rotate it to give a visual perspective.
        chart.Title.Text = "3‑D Column Chart";
        chart.Title.Rotation = 45; // rotate title 45 degrees

        // Ensure that 3‑D effects are rendered with full quality.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Dml3DEffectsRenderingMode = Dml3DEffectsRenderingMode.Advanced
        };

        // Save the document.
        doc.Save("3DColumnChart.pdf", saveOptions);
    }
}
