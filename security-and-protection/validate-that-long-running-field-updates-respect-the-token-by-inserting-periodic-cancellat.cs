using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    // Path where sample files will be stored.
    private static readonly string ArtifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");

    public static void Main()
    {
        // Ensure the output directory exists.
        Directory.CreateDirectory(ArtifactsDir);

        // 1. Create a document with many fields to simulate a long‑running update.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        for (int i = 0; i < 2000; i++)
        {
            // Insert a PAGE field; each field requires a layout pass when updated.
            builder.InsertField(FieldType.FieldPage, true);
            builder.Writeln(); // separate fields with a line break.
        }

        string sourcePath = Path.Combine(ArtifactsDir, "LongRunningFields.docx");
        doc.Save(sourcePath);

        // 2. Load the document back.
        Document loadedDoc = new Document(sourcePath);

        // 3. Update fields without cancellation – should complete successfully.
        UpdateFieldsWithCancellation(loadedDoc, CancellationToken.None);
        string updatedPath = Path.Combine(ArtifactsDir, "UpdatedWithoutCancellation.docx");
        loadedDoc.Save(updatedPath);
        if (!File.Exists(updatedPath))
            throw new InvalidOperationException("Document was not saved as expected.");

        // 4. Update fields with a cancellation token that will be triggered mid‑process.
        // Use a token that cancels after a short delay.
        using (CancellationTokenSource cts = new CancellationTokenSource())
        {
            // Cancel after 10 milliseconds.
            cts.CancelAfter(10);

            try
            {
                UpdateFieldsWithCancellation(loadedDoc, cts.Token);
                // If we reach this point, cancellation did not occur as expected.
                throw new InvalidOperationException("Cancellation was expected but did not occur.");
            }
            catch (OperationCanceledException)
            {
                // Expected path – the operation was cancelled.
                Console.WriteLine("Field update was correctly cancelled.");
            }
        }
    }

    /// <summary>
    /// Updates all fields in the document, checking the cancellation token periodically.
    /// Throws <see cref="OperationCanceledException"/> if the token is cancelled.
    /// </summary>
    private static void UpdateFieldsWithCancellation(Document doc, CancellationToken token)
    {
        // Iterate over a snapshot of the fields collection to avoid modification issues.
        FieldCollection fields = doc.Range.Fields;
        int checkInterval = 100; // Check cancellation every 100 fields.

        for (int i = 0; i < fields.Count; i++)
        {
            fields[i].Update();

            // Periodic cancellation check.
            if (i % checkInterval == 0 && token.IsCancellationRequested)
                token.ThrowIfCancellationRequested();
        }
    }
}
