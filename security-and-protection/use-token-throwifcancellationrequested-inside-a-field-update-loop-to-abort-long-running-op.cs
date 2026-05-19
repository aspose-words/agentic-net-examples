using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Define paths for the sample documents.
        string outputDir = Path.Combine(Path.GetTempPath(), "AsposeWordsDemo");
        Directory.CreateDirectory(outputDir);
        string sourcePath = Path.Combine(outputDir, "Source.docx");
        string resultPath = Path.Combine(outputDir, "Result.docx");

        // -----------------------------------------------------------------
        // 1. Create a sample document with many fields (e.g., PAGE fields).
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert 500 PAGE fields to simulate a long‑running update operation.
        for (int i = 0; i < 500; i++)
        {
            builder.Writeln($"Page number {i + 1}: ");
            builder.InsertField(FieldType.FieldPage, true);
            builder.Writeln(); // New line after each field.
        }

        // Save the source document.
        doc.Save(sourcePath);

        // -----------------------------------------------------------------
        // 2. Load the document and update fields with cancellation support.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(sourcePath);

        // Set up a cancellation token that will be triggered after a short delay.
        using (CancellationTokenSource cts = new CancellationTokenSource())
        {
            // Cancel after 10 milliseconds to force early termination.
            cts.CancelAfter(TimeSpan.FromMilliseconds(10));
            CancellationToken token = cts.Token;

            try
            {
                // Iterate over each field individually.
                foreach (Field field in loadedDoc.Range.Fields)
                {
                    // Abort if cancellation has been requested.
                    token.ThrowIfCancellationRequested();

                    // Update the current field.
                    field.Update();
                }
            }
            catch (OperationCanceledException)
            {
                // Expected path when the operation is cancelled.
                Console.WriteLine("Field update was cancelled.");
            }
        }

        // Save the (potentially partially updated) document.
        loadedDoc.Save(resultPath);

        // Verify that the output file exists.
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("Result document was not saved.");

        // Clean up (optional).
        // Directory.Delete(outputDir, true);
    }
}
