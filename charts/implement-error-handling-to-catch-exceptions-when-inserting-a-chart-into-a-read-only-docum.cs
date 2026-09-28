using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class ChartReadOnlyStreamExample
{
    public static void Main()
    {
        // Create a simple document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample document for chart insertion.");

        // Save the document to a writable memory stream.
        using (MemoryStream writableStream = new MemoryStream())
        {
            doc.Save(writableStream, SaveFormat.Docx);
            byte[] docBytes = writableStream.ToArray();

            // Create a read‑only memory stream from the saved bytes.
            using (MemoryStream readOnlyStream = new MemoryStream(docBytes, writable: false))
            {
                // Load the document from the read‑only stream.
                Document readOnlyDoc = new Document(readOnlyStream);

                try
                {
                    // Attempt to insert a chart into the read‑only document.
                    DocumentBuilder roBuilder = new DocumentBuilder(readOnlyDoc);
                    roBuilder.MoveToDocumentEnd();
                    Shape chartShape = roBuilder.InsertChart(ChartType.Column, 432, 252);
                    Chart chart = chartShape.Chart;
                    chart.Series.Clear();
                    chart.Series.Add("Sales", new double[] { 10, 20, 30 });

                    // Attempt to save back to the same read‑only stream.
                    // This will raise an exception because the stream is not writable.
                    readOnlyStream.Position = 0;
                    readOnlyDoc.Save(readOnlyStream, SaveFormat.Docx);
                    Console.WriteLine("Chart inserted and document saved successfully (unexpected).");
                }
                catch (Exception ex)
                {
                    // Handle the exception that occurs due to the read‑only stream.
                    Console.WriteLine("An error occurred while inserting the chart or saving the document:");
                    Console.WriteLine($"{ex.GetType().Name}: {ex.Message}");
                }
            }
        }
    }
}
