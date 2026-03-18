using System;
using Aspose.Words;
using Aspose.Words.Saving;

// Example paths – replace with actual locations as needed.
string inputPath = @"C:\Docs\Big document.docx";
string outputPath = @"C:\Docs\Result.html";

Document doc = null;
try
{
    // Load the document. If loading fails, an exception is thrown before "doc" is assigned.
    doc = new Document(inputPath);

    // Configure saving options with a progress callback that may abort the operation.
    HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
    {
        ProgressCallback = new SavingProgressCallback()
    };

    try
    {
        // Save operation follows the provided Save method rule.
        doc.Save(outputPath, saveOptions);
    }
    catch (OperationCanceledException ex)
    {
        // Handle the cancellation – the document will be disposed in the finally block.
        Console.WriteLine($"Saving was canceled: {ex.Message}");
    }
}
catch (FileCorruptedException ex)
{
    // Handle loading errors separately if needed.
    Console.WriteLine($"Failed to load document: {ex.Message}");
}
finally
{
    // Document may or may not implement IDisposable depending on the Aspose.Words version.
    // The safe‑cast ensures the code compiles for all supported versions and guarantees disposal when possible.
    if (doc is IDisposable disposable)
        disposable.Dispose();
}

// Callback that aborts saving after a short duration (uses the IDocumentSavingCallback rule).
public class SavingProgressCallback : IDocumentSavingCallback
{
    private readonly DateTime _startTime = DateTime.Now;
    private const double MaxDurationSeconds = 0.1; // Adjust as required.

    public void Notify(DocumentSavingArgs args)
    {
        double elapsed = (DateTime.Now - _startTime).TotalSeconds;
        if (elapsed > MaxDurationSeconds)
            throw new OperationCanceledException(
                $"EstimatedProgress = {args.EstimatedProgress}; CanceledAt = {DateTime.Now}");
    }
}
