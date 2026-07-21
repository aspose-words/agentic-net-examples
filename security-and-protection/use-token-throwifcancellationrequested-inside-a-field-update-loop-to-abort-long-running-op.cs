using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a cancellation token that will be triggered after a short delay.
        var cts = new CancellationTokenSource();
        cts.CancelAfter(100); // 100 milliseconds

        // Create a blank document and add many fields to simulate a long‑running update.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        for (int i = 0; i < 5000; i++)
        {
            builder.Writeln("Page number:");
            // Insert a PAGE field using the correct FieldType enum value.
            builder.InsertField(FieldType.FieldPage, true);
        }

        // Attempt to update all fields, aborting if the token is cancelled.
        try
        {
            UpdateFieldsWithCancellation(doc, cts.Token);
        }
        catch (OperationCanceledException ex)
        {
            Console.WriteLine($"Field update cancelled: {ex.Message}");
        }

        // Save the document to a temporary location.
        string outputPath = Path.Combine(Path.GetTempPath(), "AsposeFields.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Document was not saved successfully.");

        Console.WriteLine($"Document saved to: {outputPath}");
    }

    // Updates each field in the document, checking the cancellation token on each iteration.
    private static void UpdateFieldsWithCancellation(Document doc, CancellationToken token)
    {
        foreach (Field field in doc.Range.Fields)
        {
            token.ThrowIfCancellationRequested(); // Abort if cancellation was requested.
            field.Update();
        }
    }
}
