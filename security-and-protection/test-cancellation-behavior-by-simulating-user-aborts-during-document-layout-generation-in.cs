using System;
using System.IO;
using System.Diagnostics;
using Aspose.Words;
using Aspose.Words.Saving;

public class SavingProgressCallback : IDocumentSavingCallback
{
    private readonly Stopwatch _stopwatch = Stopwatch.StartNew();
    private const double MaxDuration = 0.01; // seconds

    public void Notify(DocumentSavingArgs args)
    {
        if (_stopwatch.Elapsed.TotalSeconds > MaxDuration)
            throw new OperationCanceledException($"EstimatedProgress = {args.EstimatedProgress}");
    }
}

public class Program
{
    public static void Main()
    {
        // Create a document with enough content to make layout processing noticeable.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        for (int i = 0; i < 1000; i++)
        {
            builder.Writeln($"Line {i + 1}");
        }

        string outputPath = "CanceledLayout.docx";

        // Clean up any previous output.
        if (File.Exists(outputPath))
            File.Delete(outputPath);

        // Set up save options with a progress callback that aborts quickly.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
        {
            ProgressCallback = new SavingProgressCallback()
        };

        bool canceled = false;
        try
        {
            doc.Save(outputPath, saveOptions);
        }
        catch (OperationCanceledException ex)
        {
            canceled = true;
            Console.WriteLine($"Save operation was canceled: {ex.Message}");
        }

        // Verify that the operation was indeed canceled.
        if (!canceled)
            throw new Exception("Expected the save operation to be canceled, but it completed.");

        // Verify that no output file was created.
        if (File.Exists(outputPath))
            throw new Exception("Output file should not exist after cancellation.");

        Console.WriteLine("Cancellation behavior test passed.");
    }
}
