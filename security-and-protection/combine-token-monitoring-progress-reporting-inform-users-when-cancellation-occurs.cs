using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Path to the source document.
        string sourcePath = @"MyDir\Big document.docx";
        // Path to the output document.
        string outputPath = @"ArtifactsDir\ProcessedDocument.docx";

        // Set up loading progress callback that reports progress and cancels after a time limit.
        var loadingCallback = new LoadingProgressCallback();

        // LoadOptions with the progress callback.
        LoadOptions loadOptions = new LoadOptions { ProgressCallback = loadingCallback };

        Document doc = null;
        try
        {
            // Load the document using the provided LoadOptions (creation rule).
            doc = new Document(sourcePath, loadOptions);
        }
        catch (OperationCanceledException ex)
        {
            Console.WriteLine($"Loading canceled: {ex.Message}");
            return;
        }

        // Attach field updating callbacks to monitor field processing.
        var fieldCallback = new FieldUpdatingCallback();
        doc.FieldOptions.FieldUpdatingCallback = fieldCallback;               // IFieldUpdatingCallback
        doc.FieldOptions.FieldUpdatingProgressCallback = fieldCallback;      // IFieldUpdatingProgressCallback

        try
        {
            // Update fields; progress will be reported via the callback.
            doc.UpdateFields();
        }
        catch (OperationCanceledException ex)
        {
            Console.WriteLine($"Field updating canceled: {ex.Message}");
            return;
        }

        // Set up saving progress callback that reports progress and cancels after a time limit.
        var savingCallback = new SavingProgressCallback();

        // SaveOptions is abstract – create a concrete instance for the desired format (DOCX here).
        SaveOptions saveOptions = SaveOptions.CreateSaveOptions(SaveFormat.Docx);
        saveOptions.ProgressCallback = savingCallback;

        try
        {
            // Save the document using the provided SaveOptions (saving rule).
            doc.Save(outputPath, saveOptions);
        }
        catch (OperationCanceledException ex)
        {
            Console.WriteLine($"Saving canceled: {ex.Message}");
        }
    }
}

// Loading progress callback implementation.
public class LoadingProgressCallback : IDocumentLoadingCallback
{
    private readonly DateTime _startTime;
    private const double MaxDurationSeconds = 0.5; // Adjust as needed.

    public LoadingProgressCallback()
    {
        _startTime = DateTime.Now;
    }

    public void Notify(DocumentLoadingArgs args)
    {
        double elapsed = (DateTime.Now - _startTime).TotalSeconds;
        Console.WriteLine($"Loading progress: {args.EstimatedProgress:F2}% (elapsed {elapsed:F2}s)");

        if (elapsed > MaxDurationSeconds)
            throw new OperationCanceledException($"Loading canceled at {elapsed:F2}s, progress {args.EstimatedProgress:F2}%");
    }
}

// Saving progress callback implementation.
public class SavingProgressCallback : IDocumentSavingCallback
{
    private readonly DateTime _startTime;
    private const double MaxDurationSeconds = 0.1; // Adjust as needed.

    public SavingProgressCallback()
    {
        _startTime = DateTime.Now;
    }

    public void Notify(DocumentSavingArgs args)
    {
        double elapsed = (DateTime.Now - _startTime).TotalSeconds;
        Console.WriteLine($"Saving progress: {args.EstimatedProgress:F2}% (elapsed {elapsed:F2}s)");

        if (elapsed > MaxDurationSeconds)
            throw new OperationCanceledException($"Saving canceled at {elapsed:F2}s, progress {args.EstimatedProgress:F2}%");
    }
}

// Field updating callback that also reports progress.
public class FieldUpdatingCallback : IFieldUpdatingCallback, IFieldUpdatingProgressCallback
{
    private readonly DateTime _startTime;
    private const double MaxDurationSeconds = 0.3; // Adjust as needed.

    public FieldUpdatingCallback()
    {
        _startTime = DateTime.Now;
        UpdatedFields = new List<string>();
    }

    // Called before a field is updated.
    void IFieldUpdatingCallback.FieldUpdating(Field field)
    {
        // Example: modify author field before update.
        if (field.Type == FieldType.FieldAuthor)
        {
            var authorField = (FieldAuthor)field;
            authorField.AuthorName = "Updated Author";
        }
    }

    // Called after a field is updated.
    void IFieldUpdatingCallback.FieldUpdated(Field field)
    {
        UpdatedFields.Add(field.Result);
    }

    // Called to report progress of field updating.
    public void Notify(FieldUpdatingProgressArgs args)
    {
        double elapsed = (DateTime.Now - _startTime).TotalSeconds;
        Console.WriteLine($"Field updating: {args.UpdatedFieldsCount}/{args.TotalFieldsCount} (completed: {args.UpdateCompleted}) (elapsed {elapsed:F2}s)");

        if (elapsed > MaxDurationSeconds)
            throw new OperationCanceledException($"Field updating canceled after {elapsed:F2}s, {args.UpdatedFieldsCount}/{args.TotalFieldsCount} fields processed.");
    }

    public IList<string> UpdatedFields { get; }
}
