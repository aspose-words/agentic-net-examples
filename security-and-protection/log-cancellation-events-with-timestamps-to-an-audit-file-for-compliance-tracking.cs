using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Saving;

public class Program
{
    // Callback that cancels the save operation after a short duration.
    private class SavingProgressCallback : IDocumentSavingCallback
    {
        private readonly DateTime _startTime;
        private const double MaxDurationSeconds = 0.001; // Very short to trigger cancellation.

        public SavingProgressCallback()
        {
            _startTime = DateTime.Now;
        }

        public void Notify(DocumentSavingArgs args)
        {
            double elapsed = (DateTime.Now - _startTime).TotalSeconds;
            if (elapsed > MaxDurationSeconds)
                throw new OperationCanceledException(
                    $"Save operation canceled after {elapsed:F4} seconds. EstimatedProgress={args.EstimatedProgress}");
        }
    }

    public static void Main()
    {
        // Prepare a simple document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample content for cancellation test.");

        // Configure save options with the progress callback.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
        {
            ProgressCallback = new SavingProgressCallback()
        };

        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");
        string auditPath = Path.Combine(Directory.GetCurrentDirectory(), "audit.log");

        try
        {
            // Attempt to save; this should be canceled by the callback.
            doc.Save(outputPath, saveOptions);
        }
        catch (OperationCanceledException ex)
        {
            // Log the cancellation event with a timestamp.
            string logEntry = $"{DateTime.UtcNow:O} - Cancellation event: {ex.Message}";
            File.AppendAllText(auditPath, logEntry + Environment.NewLine);
        }

        // Verify that the audit file was created (optional validation).
        if (!File.Exists(auditPath))
            throw new InvalidOperationException("Audit log was not created.");

        // End of program – no interactive prompts.
    }
}
