using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    // Path for the audit log file.
    private const string AuditFilePath = "audit_log.txt";

    // Simple implementation of IDocumentSavingCallback that logs cancellation events.
    private class SavingProgressCallback : IDocumentSavingCallback
    {
        private readonly DateTime _startTime;
        private readonly TimeSpan _maxDuration;

        public SavingProgressCallback(TimeSpan maxDuration)
        {
            _startTime = DateTime.UtcNow;
            _maxDuration = maxDuration;
        }

        public void Notify(DocumentSavingArgs args)
        {
            // If the saving operation exceeds the allowed duration, log and abort.
            if (DateTime.UtcNow - _startTime > _maxDuration)
            {
                string logEntry = $"{DateTime.UtcNow:O} - Document save cancelled. EstimatedProgress={args.EstimatedProgress:F2}%{Environment.NewLine}";
                File.AppendAllText(AuditFilePath, logEntry);
                throw new OperationCanceledException("Saving operation exceeded the maximum allowed duration.");
            }
        }
    }

    public static void Main()
    {
        // Ensure the audit file exists.
        if (!File.Exists(AuditFilePath))
            File.WriteAllText(AuditFilePath, $"Audit Log Started at {DateTime.UtcNow:O}{Environment.NewLine}");

        // Create a simple Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document used to demonstrate cancellation logging.");

        // Configure save options with a progress callback that will cancel after 0.1 seconds.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
        {
            ProgressCallback = new SavingProgressCallback(TimeSpan.FromMilliseconds(100))
        };

        try
        {
            // Attempt to save the document. The callback is expected to cancel the operation.
            doc.Save("sample_output.docx", saveOptions);
        }
        catch (OperationCanceledException ex)
        {
            // Log the exception details as part of the audit.
            string logEntry = $"{DateTime.UtcNow:O} - Save operation aborted: {ex.Message}{Environment.NewLine}";
            File.AppendAllText(AuditFilePath, logEntry);
        }

        // Indicate completion (no interactive output required).
    }
}
