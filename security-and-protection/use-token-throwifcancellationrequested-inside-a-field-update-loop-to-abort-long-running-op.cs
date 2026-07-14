using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a blank document and add a large number of fields to simulate a long‑running update.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        const int fieldCount = 5000;
        for (int i = 0; i < fieldCount; i++)
        {
            builder.Writeln($"Paragraph before field {i}");
            // Insert a simple MERGEFIELD; each field requires processing when updated.
            builder.InsertField($"MERGEFIELD Field{i}");
            builder.Writeln($"Paragraph after field {i}");
        }

        // Set up a cancellation token that will trigger after a short delay.
        using (CancellationTokenSource cts = new CancellationTokenSource())
        {
            // Cancel after 100 milliseconds – enough to interrupt the update loop.
            cts.CancelAfter(100);

            try
            {
                // Iterate over all fields and update them one by one.
                // ThrowIfCancellationRequested is called inside the loop to abort efficiently.
                foreach (Field field in doc.Range.Fields)
                {
                    cts.Token.ThrowIfCancellationRequested();
                    field.Update();
                }

                Console.WriteLine("All fields updated successfully.");
            }
            catch (OperationCanceledException)
            {
                Console.WriteLine("Field update operation was cancelled.");
            }
        }

        // Save the (partially) updated document to the local file system.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "UpdatedDocument.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
