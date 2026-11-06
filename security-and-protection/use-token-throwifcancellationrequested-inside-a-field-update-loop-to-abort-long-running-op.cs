using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Fields;   // Needed for the Field class

public class Program
{
    public static void Main()
    {
        // Create a sample document with many fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        for (int i = 0; i < 5000; i++)
        {
            builder.Writeln($"Field {i}: ");
            builder.InsertField("MERGEFIELD FieldName", $"Value{i}");
        }

        // Save the initial document.
        string sourcePath = "Source.docx";
        doc.Save(sourcePath);

        // Prepare a cancellation token source.
        CancellationTokenSource cts = new CancellationTokenSource();

        // Simulate cancellation after a certain number of fields have been processed.
        int cancelAfter = 1000;
        int processed = 0;

        try
        {
            foreach (Field field in doc.Range.Fields)
            {
                // Abort if cancellation has been requested.
                cts.Token.ThrowIfCancellationRequested();

                // Update the current field.
                field.Update();

                processed++;
                if (processed == cancelAfter)
                {
                    // Trigger cancellation.
                    cts.Cancel();
                }
            }
        }
        catch (OperationCanceledException)
        {
            // Handle the cancellation gracefully.
            Console.WriteLine($"Field update cancelled after processing {processed} fields.");
        }

        // Save the (partially) updated document.
        string outputPath = "Updated.docx";
        doc.Save(outputPath);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The updated document was not saved.");

        // Program completed.
        Console.WriteLine("Processing completed.");
    }
}
