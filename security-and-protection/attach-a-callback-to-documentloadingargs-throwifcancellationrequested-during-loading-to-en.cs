using System;
using System.Threading;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a simple document and save it locally.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, Aspose.Words!");
        const string fileName = "sample.docx";
        doc.Save(fileName);

        // Prepare a cancelled token to simulate a cancellation request.
        var cts = new CancellationTokenSource();
        cts.Cancel(); // The token is now in the cancelled state.

        try
        {
            // Check for cancellation before attempting to load the document.
            // This mimics the behaviour of a loading callback that would call
            // DocumentLoadingArgs.ThrowIfCancellationRequested().
            if (cts.Token.IsCancellationRequested)
                throw new OperationCanceledException();

            // Load the document (this line will not be reached because of the cancellation above).
            var loadedDoc = new Document(fileName);
            Console.WriteLine("Document loaded successfully (unexpected).");
        }
        catch (OperationCanceledException)
        {
            Console.WriteLine("Document loading was cancelled as expected.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected exception occurred: {ex.Message}");
        }
    }
}
