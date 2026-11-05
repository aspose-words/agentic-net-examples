using System;
using System.Threading;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        Document doc = null;
        try
        {
            // Create a new document and add simple content.
            doc = new Document();
            var builder = new DocumentBuilder(doc);
            builder.Writeln("Sample text.");

            // Perform an operation that may be cancelled.
            PerformCancellableOperation();
        }
        catch (OperationCanceledException)
        {
            // Handle the cancellation.
            Console.WriteLine("Operation was cancelled.");
        }
        finally
        {
            // Aspose.Words.Document does not implement IDisposable, so no explicit Dispose is required.
            // Setting the reference to null allows the garbage collector to reclaim the object.
            doc = null;
        }

        Console.WriteLine("Program completed.");
    }

    private static void PerformCancellableOperation()
    {
        // Simulate a cancellation scenario.
        throw new OperationCanceledException("Simulated cancellation.");
    }
}
