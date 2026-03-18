using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

namespace AsposeWordsCancellationDemo
{
    // Simple data source for the reporting engine.
    public class ReportData
    {
        public string Title { get; set; }
        public int Value { get; set; }
    }

    // Callback that aborts document loading after a short time.
    public class LoadingProgressCallback : IDocumentLoadingCallback
    {
        private readonly DateTime _start;
        private const double MaxDurationSeconds = 0.05; // Cancel after 50 ms.

        public LoadingProgressCallback()
        {
            _start = DateTime.Now;
        }

        public void Notify(DocumentLoadingArgs args)
        {
            double elapsed = (DateTime.Now - _start).TotalSeconds;
            if (elapsed > MaxDurationSeconds)
                throw new OperationCanceledException(
                    $"EstimatedProgress = {args.EstimatedProgress}; CanceledAt = {DateTime.Now}");
        }
    }

    // Callback that aborts document saving after a short time.
    public class SavingProgressCallback : IDocumentSavingCallback
    {
        private readonly DateTime _start;
        private const double MaxDurationSeconds = 0.05; // Cancel after 50 ms.

        public SavingProgressCallback()
        {
            _start = DateTime.Now;
        }

        public void Notify(DocumentSavingArgs args)
        {
            double elapsed = (DateTime.Now - _start).TotalSeconds;
            if (elapsed > MaxDurationSeconds)
                throw new OperationCanceledException(
                    $"EstimatedProgress = {args.EstimatedProgress}; CanceledAt = {DateTime.Now}");
        }
    }

    class Program
    {
        static void Main()
        {
            // Paths to the template and the output document.
            string templatePath = "Template.docx";
            string outputPath = "Report_Output.docx";

            // -----------------------------------------------------------------
            // 1. Load the template with a loading progress callback that cancels.
            // -----------------------------------------------------------------
            Document document = null;
            var loadOptions = new LoadOptions
            {
                ProgressCallback = new LoadingProgressCallback()
            };

            try
            {
                document = new Document(templatePath, loadOptions);
                Console.WriteLine("Document loaded without cancellation (unexpected).");
            }
            catch (OperationCanceledException ex)
            {
                Console.WriteLine($"Loading canceled as expected: {ex.Message}");
            }

            // If loading was cancelled, load again without the callback to continue the demo.
            if (document == null)
            {
                document = new Document(templatePath);
                Console.WriteLine("Document reloaded without cancellation for further processing.");
            }

            // -----------------------------------------------------------------
            // 2. Populate the template using ReportingEngine (third‑party tool).
            // -----------------------------------------------------------------
            var data = new ReportData
            {
                Title = "Sample Report",
                Value = 123
            };

            var engine = new ReportingEngine();
            engine.BuildReport(document, data);
            Console.WriteLine("ReportingEngine populated the document.");

            // -----------------------------------------------------------------
            // 3. Save the populated document with a saving progress callback that cancels.
            // -----------------------------------------------------------------
            var saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
            {
                ProgressCallback = new SavingProgressCallback()
            };

            try
            {
                document.Save(outputPath, saveOptions);
                Console.WriteLine("Document saved without cancellation (unexpected).");
            }
            catch (OperationCanceledException ex)
            {
                Console.WriteLine($"Saving canceled as expected: {ex.Message}");
            }
        }
    }
}
