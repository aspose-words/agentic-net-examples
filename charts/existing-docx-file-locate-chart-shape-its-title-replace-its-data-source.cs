using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using Aspose.Words.Tables;

class ReplaceChartDataSource
{
    static void Main()
    {
        // Create temporary paths for the input, data source, and output files.
        string tempDir = Path.GetTempPath();
        string inputPath = Path.Combine(tempDir, "InputDocument.docx");
        string newDataSourcePath = Path.Combine(tempDir, "NewChartData.xlsx");
        string outputPath = Path.Combine(tempDir, "OutputDocument.docx");

        // Ensure the input DOCX file exists. If not, create a minimal document.
        if (!File.Exists(inputPath))
        {
            var emptyDoc = new Document();
            emptyDoc.Save(inputPath);
        }

        // Ensure the new data source file exists (can be empty for this demo).
        if (!File.Exists(newDataSourcePath))
        {
            File.WriteAllBytes(newDataSourcePath, Array.Empty<byte>());
        }

        // The exact title text of the chart we want to modify.
        string targetChartTitle = "Sales Overview";

        // Load the existing Word document.
        Document doc = new Document(inputPath);

        // Iterate through all Shape nodes in the document (including those inside headers/footers).
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
        {
            // Ensure the shape actually contains a chart.
            Chart chart = shape.Chart;
            if (chart == null)
                continue;

            // Check if the chart's title matches the target title.
            if (chart.Title != null && chart.Title.Text == targetChartTitle)
            {
                // Replace the chart's linked data source with the new Excel file.
                chart.SourceFullName = newDataSourcePath;

                // Break after the first match (assuming titles are unique).
                break;
            }
        }

        // Save the modified document to a new file.
        doc.Save(outputPath);

        Console.WriteLine($"Processing complete. Output saved to: {outputPath}");
    }
}
