using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class ChartReadOnlyStreamExample
{
    public static void Main()
    {
        // Step 1: Create a new document and insert a chart.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
        Chart chart = chartShape.Chart;

        // Optional: customize the chart (e.g., set a title).
        chart.Title.Text = "Sample Chart";
        chart.Title.Show = true;

        // Step 2: Save the document to a writable memory stream and obtain the byte array.
        byte[] docBytes;
        using (MemoryStream writableStream = new MemoryStream())
        {
            doc.Save(writableStream, SaveFormat.Docx);
            docBytes = writableStream.ToArray();
        }

        // Step 3: Create a read‑only memory stream from the byte array.
        // The second parameter (writable: false) makes the stream read‑only.
        using (MemoryStream readOnlyStream = new MemoryStream(docBytes, writable: false))
        {
            // Load the document from the read‑only stream.
            Document readOnlyDoc = new Document(readOnlyStream);
            DocumentBuilder readOnlyBuilder = new DocumentBuilder(readOnlyDoc);

            try
            {
                // Attempt to insert another chart into the document.
                Shape newChartShape = readOnlyBuilder.InsertChart(ChartType.Pie, 300, 300);
                Chart newChart = newChartShape.Chart;
                newChart.Title.Text = "Additional Chart";
                newChart.Title.Show = true;

                // Attempt to save the modified document back to the same read‑only stream.
                // This operation will throw because the stream does not support writing.
                readOnlyDoc.Save(readOnlyStream, SaveFormat.Docx);
                Console.WriteLine("Document saved successfully (unexpected).");
            }
            catch (Exception ex)
            {
                // Catch and display the exception that occurs due to the read‑only stream.
                Console.WriteLine("An error occurred while saving to a read‑only stream:");
                Console.WriteLine(ex.GetType().Name + ": " + ex.Message);
            }
        }

        // Step 4: Save the original (unmodified) document to a file for verification.
        doc.Save("output.docx");
        Console.WriteLine("Original document saved as 'output.docx'.");
    }
}
