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
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string sourcePath = Path.Combine(artifactsDir, "LongRunningFields.docx");
        string resultPath = Path.Combine(artifactsDir, "LongRunningFields_Updated.docx");

        // -----------------------------------------------------------------
        // 1. Create a sample document containing many fields to simulate a
        //    long‑running update operation.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Document with many PAGE fields:");
        for (int i = 0; i < 2000; i++)
        {
            // Insert a PAGE field and a line break.
            builder.InsertField(FieldType.FieldPage, true);
            builder.Writeln($" - field #{i + 1}");
        }
        doc.Save(sourcePath);

        // -----------------------------------------------------------------
        // 2. Load the document and prepare a cancellation token that will be
        //    triggered after a short timeout.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(sourcePath);
        using var cts = new CancellationTokenSource();
        // Cancel after 10 milliseconds to force early termination.
        cts.CancelAfter(TimeSpan.FromMilliseconds(10));
        CancellationToken token = cts.Token;

        bool cancelled = false;

        try
        {
            // -----------------------------------------------------------------
            // 3. Manually update each field, checking the token periodically.
            //    This mimics a long‑running field update that respects cancellation.
            // -----------------------------------------------------------------
            foreach (Field field in loadedDoc.Range.Fields)
            {
                // Throw if cancellation was requested.
                if (token.IsCancellationRequested)
                {
                    cancelled = true;
                    throw new OperationCanceledException("Field update was cancelled.");
                }

                // Update the current field.
                field.Update();
            }
        }
        catch (OperationCanceledException ex)
        {
            Console.WriteLine($"Update cancelled: {ex.Message}");
        }

        // -----------------------------------------------------------------
        // 4. Verify the outcome.
        //    If the operation was not cancelled, save the fully updated document.
        //    Otherwise, save the partially updated document for inspection.
        // -----------------------------------------------------------------
        if (!cancelled)
        {
            loadedDoc.Save(resultPath);
            Console.WriteLine($"All fields updated successfully. Saved to: {resultPath}");
        }
        else
        {
            // Save the partially updated document to demonstrate that cancellation stopped the process.
            loadedDoc.Save(resultPath);
            Console.WriteLine($"Partial update saved to: {resultPath}");
        }

        // Simple validation: ensure the output file exists.
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("The expected output document was not created.");
    }
}
