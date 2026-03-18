using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;            // <-- added
using Aspose.Words.Drawing.Charts;

class InsertChartIntoReadOnlyStream
{
    static void Main()
    {
        // Path to the source document.
        const string sourcePath = @"C:\Docs\SourceDocument.docx";
        // Path where the modified document will be saved (if insertion succeeds).
        const string outputPath = @"C:\Docs\ModifiedDocument.docx";

        // Open the source document as a read‑only stream.
        // The stream is opened with FileAccess.Read, which makes it read‑only.
        using (FileStream readOnlyStream = new FileStream(sourcePath, FileMode.Open, FileAccess.Read))
        {
            // Load the document from the read‑only stream.
            Document doc = new Document(readOnlyStream);

            // Create a DocumentBuilder for editing the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            try
            {
                // Attempt to insert a chart. This operation may throw if the document
                // cannot be modified because it was loaded from a read‑only source.
                Shape chartShape = builder.InsertChart(ChartType.Pie, 300, 300);

                // Optional: customize the chart (e.g., add a series).
                Chart chart = chartShape.Chart;
                chart.Series.Clear();
                chart.Series.Add("Sample", new[] { "A", "B", "C" }, new[] { 10.0, 20.0, 30.0 });

                // Save the modified document.
                doc.Save(outputPath);
                Console.WriteLine("Chart inserted and document saved successfully.");
            }
            catch (Exception ex)
            {
                // Catch any exception that occurs during chart insertion.
                // Aspose.Words may throw InvalidOperationException, UnauthorizedAccessException,
                // or other exceptions when the document cannot be edited.
                Console.WriteLine($"Failed to insert chart: {ex.GetType().Name} - {ex.Message}");
            }
        }
    }
}
