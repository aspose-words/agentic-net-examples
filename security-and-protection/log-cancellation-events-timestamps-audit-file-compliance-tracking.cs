using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AuditExample
{
    // Simple logger that appends timestamped messages to a file.
    class AuditLogger
    {
        private readonly string _filePath;

        public AuditLogger(string filePath)
        {
            _filePath = filePath;
        }

        public void Log(string message)
        {
            string line = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} {message}{Environment.NewLine}";
            File.AppendAllText(_filePath, line);
        }
    }

    // Callback invoked during document saving. Logs progress and cancels when the operation exceeds a time limit.
    class AuditSavingCallback : IDocumentSavingCallback
    {
        private readonly AuditLogger _logger;
        private readonly DateTime _startTime;
        private readonly double _maxDurationSeconds;

        public AuditSavingCallback(AuditLogger logger, double maxDurationSeconds = 0.1)
        {
            _logger = logger;
            _startTime = DateTime.Now;
            _maxDurationSeconds = maxDurationSeconds;
        }

        public void Notify(DocumentSavingArgs args)
        {
            double elapsed = (DateTime.Now - _startTime).TotalSeconds;

            // Log current progress.
            _logger.Log($"Saving progress: {args.EstimatedProgress:P2}");

            // If the operation runs too long, log cancellation and abort.
            if (elapsed > _maxDurationSeconds)
            {
                _logger.Log($"Saving cancelled after {elapsed:F2}s (estimated progress {args.EstimatedProgress:P2})");
                throw new OperationCanceledException($"Saving aborted at {DateTime.Now}");
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the audit log file.
            var auditLogPath = "audit.log";
            var logger = new AuditLogger(auditLogPath);

            // Create a new document and add some content.
            var doc = new Document();
            var builder = new DocumentBuilder(doc);
            builder.Writeln("This is a sample document for audit logging.");

            // Configure saving options with the custom progress callback.
            var saveOptions = new PdfSaveOptions
            {
                ProgressCallback = new AuditSavingCallback(logger)
            };

            try
            {
                // Attempt to save the document. The callback will log progress and may cancel the operation.
                doc.Save("output.pdf", saveOptions);
                logger.Log("Document saved successfully.");
            }
            catch (OperationCanceledException ex)
            {
                // Log the cancellation exception details.
                logger.Log($"OperationCanceledException caught: {ex.Message}");
            }
        }
    }
}
