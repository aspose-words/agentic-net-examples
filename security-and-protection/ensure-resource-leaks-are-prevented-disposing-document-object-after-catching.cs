using System;
using Aspose.Words;
using Aspose.Words.Saving;

string outputPath = "Result.html";

Document doc = null;
try
{
    // Create a simple document in memory.
    doc = new Document();
    var builder = new DocumentBuilder(doc);
    builder.Writeln("Hello, world!");

    var saveOptions = new HtmlSaveOptions(SaveFormat.Html)
    {
        ProgressCallback = new SavingProgressCallback()
    };

    try
    {
        doc.Save(outputPath, saveOptions);
        Console.WriteLine($"Document saved to {outputPath}");
    }
    catch (OperationCanceledException ex)
    {
        Console.WriteLine($"Saving was canceled: {ex.Message}");
    }
}
catch (FileCorruptedException ex)
{
    Console.WriteLine($"Failed to load document: {ex.Message}");
}
finally
{
    if (doc is IDisposable disposable)
        disposable.Dispose();
}

public class SavingProgressCallback : IDocumentSavingCallback
{
    private readonly DateTime _startTime = DateTime.Now;
    private const double MaxDurationSeconds = 0.1;

    public void Notify(DocumentSavingArgs args)
    {
        if ((DateTime.Now - _startTime).TotalSeconds > MaxDurationSeconds)
            throw new OperationCanceledException($"EstimatedProgress = {args.EstimatedProgress}");
    }
}
