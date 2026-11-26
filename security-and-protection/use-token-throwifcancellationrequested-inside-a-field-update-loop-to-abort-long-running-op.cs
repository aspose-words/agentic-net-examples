using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a sample document with many fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        for (int i = 0; i < 500; i++)
        {
            builder.Writeln($"Field {i}:");
            builder.InsertField("MERGEFIELD FieldName", $"Value{i}");
        }

        // Save the source document to a temporary location.
        string sourcePath = Path.Combine(Path.GetTempPath(), "Source.docx");
        doc.Save(sourcePath);

        // Prepare a cancellation token that will cancel after a short delay.
        using CancellationTokenSource cts = new CancellationTokenSource();
        cts.CancelAfter(10); // Cancel after 10 milliseconds.
        CancellationToken token = cts.Token;

        // Load the document again to simulate a real processing scenario.
        Document loadDoc = new Document(sourcePath);

        try
        {
            // Update each field, checking for cancellation on each iteration.
            foreach (Field field in loadDoc.Range.Fields)
            {
                token.ThrowIfCancellationRequested();
                field.Update();
            }
        }
        catch (OperationCanceledException)
        {
            // Expected when the operation is cancelled.
            Console.WriteLine("Field update operation was cancelled.");
        }

        // Save the (potentially partially updated) document.
        string outputPath = Path.Combine(Path.GetTempPath(), "Output.docx");
        loadDoc.Save(outputPath);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved.");

        // Indicate that processing has finished.
        Console.WriteLine("Processing completed.");
    }
}
