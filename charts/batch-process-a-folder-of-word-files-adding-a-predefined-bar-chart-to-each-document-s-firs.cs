using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;               // Needed for the Shape class
using Aspose.Words.Drawing.Charts;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Define a working directory inside the current folder.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "ChartsBatch");
        Directory.CreateDirectory(workDir);

        // Create a few sample DOCX files if the folder is empty.
        // This ensures the example is self‑contained.
        if (Directory.GetFiles(workDir, "*.docx").Length == 0)
        {
            for (int i = 1; i <= 3; i++)
            {
                Document sample = new Document();
                sample.Save(Path.Combine(workDir, $"input-{i}.docx"));
            }
        }

        // Predefined chart data.
        string[] categories = { "Q1", "Q2", "Q3", "Q4" };
        double[] values = { 15.0, 30.5, 22.0, 40.0 };
        string seriesName = "Quarterly Sales";

        // Process each DOCX file in the folder.
        foreach (string filePath in Directory.GetFiles(workDir, "*.docx"))
        {
            // Load the document.
            Document doc = new Document(filePath);
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Move the cursor to the start of the document (first page).
            builder.MoveToDocumentStart();

            // Insert a Bar chart.
            Shape chartShape = builder.InsertChart(ChartType.Bar, 432, 252);

            // Ensure the shape actually contains a chart.
            if (chartShape.HasChart)
            {
                Chart chart = chartShape.Chart;

                // Remove the demo data that Aspose.Words inserts by default.
                chart.Series.Clear();

                // Add our predefined series.
                chart.Series.Add(seriesName, categories, values);

                // Optional: set a title for the chart.
                chart.Title.Text = "Quarterly Sales Overview";
                chart.Title.Show = true;
                chart.Title.Font.Size = 14;
                chart.Title.Font.Color = Color.DarkBlue;
            }

            // Save the modified document back to the same file.
            doc.Save(filePath);
        }
    }
}
