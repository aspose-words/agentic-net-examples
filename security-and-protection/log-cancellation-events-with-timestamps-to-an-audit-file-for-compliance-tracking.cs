using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample content for cancellation test.");

        // Configure save options with a progress callback that will cancel the operation.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx);
        saveOptions.ProgressCallback = new SavingProgressCallback();

        string outputPath = "output.docx";
        string auditPath = "audit_log.txt";

        try
        {
            // Attempt to save the document; the callback will cancel it.
            doc.Save(outputPath, saveOptions);
        }
        catch (OperationCanceledException ex)
        {
            // Log the cancellation event with a timestamp.
            string logEntry = $"{DateTime.Now:o} - Save operation canceled: {ex.Message}";
            File.AppendAllText(auditPath, logEntry + Environment.NewLine);
        }

        // Ensure the audit file was created.
        if (!File.Exists(auditPath))
        {
            throw new Exception("Audit log file was not created.");
        }
    }

    // Callback that aborts the save after a short duration.
    private class SavingProgressCallback : IDocumentSavingCallback
    {
        private readonly DateTime _start = DateTime.Now;
        private const double MaxDurationSeconds = 0.001; // Cancel after ~1 ms.

        public void Notify(DocumentSavingArgs args)
        {
            double elapsed = (DateTime.Now - _start).TotalSeconds;
            if (elapsed > MaxDurationSeconds)
                throw new OperationCanceledException($"EstimatedProgress = {args.EstimatedProgress}");
        }
    }
}
