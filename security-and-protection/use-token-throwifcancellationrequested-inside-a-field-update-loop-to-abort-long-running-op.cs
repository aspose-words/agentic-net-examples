using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a sample document with many PAGE fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        for (int i = 0; i < 1000; i++)
        {
            builder.Writeln($"Page {i + 1}: ");
            builder.InsertField(FieldType.FieldPage, true);
        }

        // Save the original document.
        string originalPath = "Original.docx";
        doc.Save(originalPath);
        if (!File.Exists(originalPath))
            throw new InvalidOperationException("Failed to create the original document.");

        // Prepare a cancellation token that will be triggered shortly.
        using (CancellationTokenSource cts = new CancellationTokenSource())
        {
            // Cancel after 10 milliseconds to simulate a timeout.
            cts.CancelAfter(10);
            CancellationToken token = cts.Token;

            try
            {
                // Loop through all fields and update them, checking for cancellation.
                foreach (Field field in doc.Range.Fields)
                {
                    token.ThrowIfCancellationRequested();
                    field.Update();
                }
            }
            catch (OperationCanceledException)
            {
                Console.WriteLine("Field update operation was cancelled.");
            }
        }

        // Save the (potentially partially updated) document.
        string updatedPath = "Updated.docx";
        doc.Save(updatedPath);
        if (!File.Exists(updatedPath))
            throw new InvalidOperationException("Failed to save the updated document.");

        // Indicate completion.
        Console.WriteLine("Processing completed.");
    }
}
