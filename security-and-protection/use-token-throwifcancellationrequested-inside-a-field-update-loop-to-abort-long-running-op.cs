using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Fields;   // Needed for the Field class

public class Program
{
    public static void Main()
    {
        // Create a sample document with many PAGE fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        for (int i = 0; i < 1000; i++)
        {
            // Insert a PAGE field with a placeholder result.
            builder.InsertField("PAGE", $"Page {i + 1}");
            builder.Writeln();
        }

        // Define where the updated document will be saved.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "UpdatedFields.docx");

        // Set up a cancellation token that triggers after a short delay.
        using (CancellationTokenSource cts = new CancellationTokenSource())
        {
            cts.CancelAfter(100); // Cancel after 100 ms.
            CancellationToken token = cts.Token;

            try
            {
                // Update each field, aborting if cancellation is requested.
                foreach (Field field in doc.Range.Fields)
                {
                    token.ThrowIfCancellationRequested();
                    field.Update();
                }
            }
            catch (OperationCanceledException)
            {
                Console.WriteLine("Field update was cancelled.");
            }
        }

        // Save the document (whether fully updated or partially).
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved.");

        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
