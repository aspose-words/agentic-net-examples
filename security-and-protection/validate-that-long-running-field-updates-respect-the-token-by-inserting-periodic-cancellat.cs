using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample document with many fields to simulate a long‑running update.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        for (int i = 0; i < 2000; i++)
        {
            builder.Writeln($"Paragraph {i + 1}");
            // Insert a PAGE field; updating many fields will take noticeable time.
            builder.InsertField(FieldType.FieldPage, true);
        }

        // Save the original document (optional, just to have a file on disk).
        string sourcePath = Path.Combine(outputDir, "Source.docx");
        doc.Save(sourcePath);

        // Set up a cancellation token that will be triggered after a short delay.
        using (CancellationTokenSource cts = new CancellationTokenSource())
        {
            // Cancel after 10 milliseconds.
            cts.CancelAfter(10);
            CancellationToken token = cts.Token;

            try
            {
                // Manually update each field, checking the token periodically.
                // This mimics a long‑running field update respecting cancellation.
                foreach (Field field in doc.Range.Fields)
                {
                    // Throw if cancellation was requested.
                    token.ThrowIfCancellationRequested();

                    // Update the current field.
                    field.Update();
                }

                // If we reach this point, the update completed without cancellation.
                // Save the updated document.
                string updatedPath = Path.Combine(outputDir, "Updated.docx");
                doc.Save(updatedPath);

                // Validation: the document should exist.
                if (!File.Exists(updatedPath))
                    throw new InvalidOperationException("Updated document was not saved as expected.");
            }
            catch (OperationCanceledException)
            {
                // Expected path: the operation was cancelled.
                // Save the partially updated document to verify that cancellation was respected.
                string cancelledPath = Path.Combine(outputDir, "Cancelled.docx");
                doc.Save(cancelledPath);

                // Validation: the cancelled document should exist.
                if (!File.Exists(cancelledPath))
                    throw new InvalidOperationException("Cancelled document was not saved as expected.");
            }
        }
    }
}
