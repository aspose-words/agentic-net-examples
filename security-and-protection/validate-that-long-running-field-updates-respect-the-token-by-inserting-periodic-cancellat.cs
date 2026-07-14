using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    // Updates all fields in the document while periodically checking the cancellation token.
    private static void UpdateFieldsWithCancellation(Document doc, CancellationToken token)
    {
        int processed = 0;
        foreach (Field field in doc.Range.Fields)
        {
            // Throw if cancellation was requested.
            token.ThrowIfCancellationRequested();

            field.Update();
            processed++;

            // Additional periodic check (every 100 fields) to simulate long‑running work.
            if (processed % 100 == 0)
                token.ThrowIfCancellationRequested();
        }
    }

    public static void Main()
    {
        // Prepare a temporary folder for the sample files.
        string artifactsDir = Path.Combine(Path.GetTempPath(), "AsposeWordsDemo");
        Directory.CreateDirectory(artifactsDir);
        string filePath = Path.Combine(artifactsDir, "LongRunningFields.docx");

        // Create a document with many PAGE fields to simulate a long‑running update.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        for (int i = 0; i < 5000; i++)
        {
            builder.Writeln($"Record {i + 1}:");
            // Use the FieldType overload to avoid ambiguity with older Aspose.Words versions.
            builder.InsertField(FieldType.FieldPage, true);
        }
        doc.Save(filePath);

        // Reload the document to ensure a fresh instance.
        Document loadedDoc = new Document(filePath);

        // Set up a cancellation token that will be triggered shortly after the update starts.
        using (CancellationTokenSource cts = new CancellationTokenSource())
        {
            // Cancel after a very short delay (e.g., 5 milliseconds).
            cts.CancelAfter(5);

            bool cancelled = false;
            try
            {
                UpdateFieldsWithCancellation(loadedDoc, cts.Token);
            }
            catch (OperationCanceledException)
            {
                cancelled = true;
                Console.WriteLine("Field update was cancelled as expected.");
            }

            if (!cancelled)
                throw new InvalidOperationException("The field update was expected to be cancelled but completed normally.");
        }

        // Clean up the temporary files.
        if (File.Exists(filePath))
            File.Delete(filePath);
        if (Directory.Exists(artifactsDir))
            Directory.Delete(artifactsDir, true);
    }
}
