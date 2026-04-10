using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Create a simple document in memory.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample document.");

        // Save the document to a temporary memory stream and obtain its bytes.
        using (MemoryStream tempStream = new MemoryStream())
        {
            doc.Save(tempStream, SaveFormat.Docx);
            byte[] docBytes = tempStream.ToArray();

            // Create a read‑only stream from the byte array.
            using (MemoryStream readOnlyStream = new MemoryStream(docBytes, writable: false))
            {
                // Load the document from the read‑only stream.
                Document readOnlyDoc = new Document(readOnlyStream);
                DocumentBuilder readOnlyBuilder = new DocumentBuilder(readOnlyDoc);

                // Attempt to insert a chart while the document is loaded from a read‑only stream.
                try
                {
                    Shape chartShape = readOnlyBuilder.InsertChart(ChartType.Column, 400, 300);
                    Chart chart = chartShape.Chart;
                    chart.Title.Text = "Read‑Only Test";
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error during chart insertion: {ex.Message}");
                }

                // Attempt to save the modified document back to the same read‑only stream.
                try
                {
                    readOnlyDoc.Save(readOnlyStream, SaveFormat.Docx);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error during saving to read‑only stream: {ex.Message}");
                }

                // Save the document to a regular file to verify that the chart was added.
                readOnlyDoc.Save("Result.docx");
            }
        }
    }
}
