using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple document with some text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample document.");

        // Save the document into a memory buffer.
        using (MemoryStream tempStream = new MemoryStream())
        {
            // Save to the stream using a format overload (required by the API).
            doc.Save(tempStream, SaveFormat.Docx);
            byte[] docBytes = tempStream.ToArray();

            // Create a read‑only stream from the buffer.
            using (MemoryStream readOnlyStream = new MemoryStream(docBytes, writable: false))
            {
                try
                {
                    // Load the document from the read‑only stream.
                    Document readOnlyDoc = new Document(readOnlyStream);

                    // Insert a chart into the loaded document.
                    DocumentBuilder chartBuilder = new DocumentBuilder(readOnlyDoc);
                    Shape chartShape = chartBuilder.InsertChart(ChartType.Column, 432, 252);
                    Chart chart = chartShape.Chart;

                    // Replace the default demo data with custom series.
                    chart.Series.Clear();
                    chart.Series.Add(
                        "Series 1",
                        new[] { "Category A", "Category B", "Category C" },
                        new[] { 10.0, 20.0, 30.0 });

                    // Attempt to save back to the same read‑only stream.
                    // This will throw because the stream is not writable.
                    readOnlyDoc.Save(readOnlyStream, SaveFormat.Docx);
                }
                catch (Exception ex)
                {
                    // Expected exception for the read‑only stream scenario.
                    Console.WriteLine($"Exception caught: {ex.Message}");
                }
            }
        }

        // Save the original document to a file to demonstrate normal operation.
        doc.Save("output.docx");
    }
}
