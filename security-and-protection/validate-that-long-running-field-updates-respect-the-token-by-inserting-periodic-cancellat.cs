using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    // Updates all fields in the document, checking the cancellation token periodically.
    private static void UpdateFieldsWithCancellation(Document doc, CancellationToken token)
    {
        // Retrieve the collection of fields.
        FieldCollection fields = doc.Range.Fields;

        // Update each field individually.
        for (int i = 0; i < fields.Count; i++)
        {
            // Throw if cancellation was requested.
            if (token.IsCancellationRequested)
                throw new OperationCanceledException("Field update was cancelled.");

            // Update the current field.
            fields[i].Update();
        }
    }

    public static void Main()
    {
        // Prepare output directory.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a sample document with many fields to simulate a long‑running update.
        Document doc = new Document();
        var builder = new DocumentBuilder(doc);
        for (int i = 0; i < 5000; i++)
        {
            builder.Writeln($"Field {i + 1}: ");
            builder.InsertField("PAGE", null);
        }

        // Save the initial document (optional, just for inspection).
        string sourcePath = Path.Combine(artifactsDir, "Source.docx");
        doc.Save(sourcePath);

        // Set up a cancellation token that will be triggered after a short delay.
        using var cts = new CancellationTokenSource();
        Task cancelTask = Task.Run(async () =>
        {
            await Task.Delay(10); // 10 ms delay before cancelling.
            cts.Cancel();
        });

        // Attempt to update fields with cancellation support.
        try
        {
            UpdateFieldsWithCancellation(doc, cts.Token);
            // If we reach this point, cancellation was not respected – fail the validation.
            throw new Exception("Cancellation token was ignored during field updates.");
        }
        catch (OperationCanceledException)
        {
            // Expected path: the operation was cancelled.
        }

        // Save the (partially) updated document to verify that the process completed.
        string resultPath = Path.Combine(artifactsDir, "Result.docx");
        doc.Save(resultPath);

        // Verify that the output files exist.
        if (!File.Exists(sourcePath) || !File.Exists(resultPath))
            throw new Exception("Expected output files were not created.");
    }
}
