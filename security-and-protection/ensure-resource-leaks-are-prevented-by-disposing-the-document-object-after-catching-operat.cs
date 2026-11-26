using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a simple document in memory.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, Aspose.Words!");

        try
        {
            // Simulate an operation that may be cancelled.
            // Here we directly throw the exception for demonstration.
            throw new OperationCanceledException("The operation was cancelled.");
        }
        catch (OperationCanceledException ex)
        {
            // Handle the cancellation.
            Console.WriteLine($"Caught cancellation: {ex.Message}");
            // No explicit Dispose needed for Document; rely on garbage collection.
        }
        finally
        {
            // Ensure the document reference is cleared so it can be collected promptly.
            doc = null;
        }

        // Optionally, you could perform other work here.
        Console.WriteLine("Program completed.");
    }
}
