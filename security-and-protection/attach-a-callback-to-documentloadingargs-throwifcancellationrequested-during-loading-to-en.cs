using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;

public class LoadingCallback : IDocumentLoadingCallback
{
    // This method is called repeatedly while the document is being loaded.
    public void Notify(DocumentLoadingArgs args)
    {
        // The DocumentLoadingArgs class does not provide a ThrowIfCancellationRequested method.
        // To cancel loading, simply throw an OperationCanceledException based on the progress information.
        // In this demo we abort the loading as soon as the first callback is invoked.
        throw new OperationCanceledException(
            $"Loading cancelled at {args.EstimatedProgress}% progress.");
    }
}

public class Program
{
    public static void Main()
    {
        // Prepare a folder for temporary files.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);

        // Create a simple document and save it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello Aspose.Words!");
        string filePath = Path.Combine(outputDir, "Sample.docx");
        doc.Save(filePath);

        // Set up load options with the custom progress callback.
        LoadOptions loadOptions = new LoadOptions
        {
            ProgressCallback = new LoadingCallback()
        };

        // Attempt to load the document; the callback will cancel the operation.
        try
        {
            Document loadedDoc = new Document(filePath, loadOptions);
            // If loading succeeds (which it shouldn't), write a message.
            Console.WriteLine("Document loaded successfully (unexpected).");
        }
        catch (OperationCanceledException ex)
        {
            // Expected path: the callback aborts loading.
            Console.WriteLine($"Loading was cancelled as intended: {ex.Message}");
        }
        catch (Exception ex)
        {
            // Any other exception is reported.
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
