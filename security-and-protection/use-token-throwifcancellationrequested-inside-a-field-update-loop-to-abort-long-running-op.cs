using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Define file paths in the current directory.
        string originalPath = Path.Combine(Environment.CurrentDirectory, "original.docx");
        string updatedPath = Path.Combine(Environment.CurrentDirectory, "updated.docx");

        // -----------------------------------------------------------------
        // 1. Create a sample document with many PAGE fields.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert 1000 PAGE fields to simulate a long‑running update operation.
        for (int i = 0; i < 1000; i++)
        {
            builder.Writeln($"Page {i + 1}: ");
            builder.InsertField(FieldType.FieldPage, true);
        }

        // Save the source document.
        doc.Save(originalPath);

        // -----------------------------------------------------------------
        // 2. Prepare a cancellation token that will trigger after a short delay.
        // -----------------------------------------------------------------
        using (CancellationTokenSource cts = new CancellationTokenSource())
        {
            // Cancel after 5 milliseconds to force early termination.
            cts.CancelAfter(TimeSpan.FromMilliseconds(5));
            CancellationToken token = cts.Token;

            // Load the document we just saved.
            Document loadedDoc = new Document(originalPath);

            try
            {
                // Iterate over all fields and update them, checking for cancellation.
                foreach (Field field in loadedDoc.Range.Fields)
                {
                    // Throw if cancellation has been requested.
                    token.ThrowIfCancellationRequested();

                    // Update the current field.
                    field.Update();
                }

                // If the loop completes without cancellation, save the fully updated document.
                loadedDoc.Save(updatedPath);
                Console.WriteLine("Document updated and saved successfully.");
            }
            catch (OperationCanceledException)
            {
                // Save the partially updated document to demonstrate abort handling.
                loadedDoc.Save(updatedPath);
                Console.WriteLine("Operation was canceled. Partial document saved.");
            }
        }

        // -----------------------------------------------------------------
        // 3. Verify that the output file exists.
        // -----------------------------------------------------------------
        if (!File.Exists(updatedPath))
        {
            throw new InvalidOperationException("The updated document was not saved as expected.");
        }
    }
}
