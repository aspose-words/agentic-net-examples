using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a temporary document with many fields to simulate a long‑running update.
        string docPath = Path.Combine(Path.GetTempPath(), "LongRunningFields.docx");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Insert 5000 PAGE fields (adjust count if needed for longer processing).
        for (int i = 0; i < 5000; i++)
        {
            builder.Writeln($"Page field {i + 1}: ");
            builder.InsertField(FieldType.FieldPage, true);
        }

        // Save the source document.
        doc.Save(docPath);

        // Prepare a cancellation token that will be triggered after a short delay.
        var cts = new CancellationTokenSource();
        Task.Run(async () =>
        {
            await Task.Delay(10); // Cancel after 10 ms.
            cts.Cancel();
        });

        // Load the document again for the update operation.
        var loadedDoc = new Document(docPath);

        // Manually update fields while checking the cancellation token periodically.
        try
        {
            var fields = loadedDoc.Range.Fields;
            for (int i = 0; i < fields.Count; i++)
            {
                // Throw if cancellation has been requested.
                if (cts.Token.IsCancellationRequested)
                    throw new OperationCanceledException(cts.Token);

                // Update the current field.
                fields[i].Update();
            }

            // If we reach this point, cancellation was not observed – fail the validation.
            throw new Exception("Field update completed without respecting the cancellation token.");
        }
        catch (OperationCanceledException)
        {
            // Expected path: the operation was cancelled.
            Console.WriteLine("Field update was correctly cancelled.");
        }

        // Clean up temporary file.
        if (File.Exists(docPath))
        {
            File.Delete(docPath);
        }
    }
}
