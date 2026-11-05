using System;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a document with many fields to make the update operation lengthy.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        for (int i = 0; i < 5000; i++)
        {
            builder.Writeln($"Field {i}: ");
            builder.InsertField(FieldType.FieldPage, true);
            builder.Writeln();
        }

        // Set up a cancellation token that will be triggered shortly.
        using (CancellationTokenSource cts = new CancellationTokenSource())
        {
            cts.CancelAfter(10); // Cancel after 10 milliseconds.

            try
            {
                // Manually update fields while checking the cancellation token periodically.
                foreach (Field field in doc.Range.Fields)
                {
                    if (cts.Token.IsCancellationRequested)
                        throw new OperationCanceledException(cts.Token);

                    field.Update();
                }

                // If we reach this point, the operation ignored the cancellation token.
                throw new Exception("Field update completed without respecting the cancellation token.");
            }
            catch (OperationCanceledException)
            {
                // Expected outcome: the update was cancelled.
                Console.WriteLine("Field update was cancelled as expected.");
            }
        }
    }
}
