using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    // Simple audit logger that appends messages with timestamps to a file.
    private static void LogAudit(string message)
    {
        string logPath = Path.Combine(Directory.GetCurrentDirectory(), "audit.log");
        string entry = $"{DateTime.UtcNow:O} - {message}";
        File.AppendAllLines(logPath, new[] { entry });
    }

    // Callback that cancels the document saving operation.
    private class CancelSavingCallback : IDocumentSavingCallback
    {
        public void Notify(DocumentSavingArgs args)
        {
            // Cancel the save operation immediately.
            throw new OperationCanceledException($"Save canceled at progress {args.EstimatedProgress}%.");
        }
    }

    public static void Main()
    {
        // Create a simple Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample content for the document.");

        // Prepare save options with the cancellation callback.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
        {
            ProgressCallback = new CancelSavingCallback()
        };

        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");

        try
        {
            // Attempt to save the document; this will be canceled by the callback.
            doc.Save(outputPath, saveOptions);
        }
        catch (OperationCanceledException ex)
        {
            // Log the cancellation event with a timestamp.
            LogAudit($"Document save canceled: {ex.Message}");
        }

        // Ensure the audit log file exists; if not, create an empty one.
        string auditPath = Path.Combine(Directory.GetCurrentDirectory(), "audit.log");
        if (!File.Exists(auditPath))
        {
            File.WriteAllText(auditPath, string.Empty);
        }
    }
}
