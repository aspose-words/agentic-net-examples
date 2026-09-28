using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a document with many PAGE fields to simulate a long‑running update.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        for (int i = 0; i < 5000; i++)
        {
            builder.Writeln($"Field {i + 1}: ");
            builder.InsertField("PAGE", null);
        }

        // Save the document to a temporary .docx file (extension required for format detection).
        string tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".docx");
        try
        {
            doc.Save(tempPath);

            // Reload the document from the saved file.
            Document loadedDoc = new Document(tempPath);

            // Prepare a cancellation token that will be cancelled shortly after the update starts.
            using (CancellationTokenSource cts = new CancellationTokenSource())
            {
                // Cancel after a brief delay (e.g., 10 ms).
                cts.CancelAfter(10);

                bool wasCancelled = false;
                try
                {
                    // Manually update each field, checking the cancellation token periodically.
                    foreach (Field field in loadedDoc.Range.Fields)
                    {
                        // Throw if cancellation has been requested.
                        cts.Token.ThrowIfCancellationRequested();

                        // Update the current field.
                        field.Update();

                        // Optional small delay to make cancellation more likely during the loop.
                        Thread.Sleep(1);
                    }
                }
                catch (OperationCanceledException)
                {
                    wasCancelled = true;
                }

                // Validate that the operation respected the cancellation token.
                if (!wasCancelled)
                {
                    throw new Exception("Field update did not respect the cancellation token.");
                }
            }
        }
        finally
        {
            // Clean up the temporary file.
            if (File.Exists(tempPath))
                File.Delete(tempPath);
        }
    }
}
