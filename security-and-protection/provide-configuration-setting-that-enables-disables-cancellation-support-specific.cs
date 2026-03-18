using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

namespace AsposeWordsCancellationDemo
{
    // Simple configuration class to enable/disable cancellation for loading and saving.
    public class CancellationConfig
    {
        // When true, loading will be cancelled after MaxDuration seconds.
        public bool EnableLoadingCancellation { get; set; } = false;

        // When true, saving will be cancelled after MaxDuration seconds.
        public bool EnableSavingCancellation { get; set; } = false;
    }

    // Callback for document loading. Throws OperationCanceledException only when enabled.
    public class LoadingProgressCallback : IDocumentLoadingCallback
    {
        private readonly DateTime _loadingStartedAt = DateTime.Now;
        private readonly double _maxDurationSeconds;
        private readonly CancellationConfig _config;

        public LoadingProgressCallback(double maxDurationSeconds, CancellationConfig config)
        {
            _maxDurationSeconds = maxDurationSeconds;
            _config = config;
        }

        public void Notify(DocumentLoadingArgs args)
        {
            // If cancellation is disabled, simply return.
            if (!_config.EnableLoadingCancellation)
                return;

            double elapsedSeconds = (DateTime.Now - _loadingStartedAt).TotalSeconds;
            if (elapsedSeconds > _maxDurationSeconds)
                throw new OperationCanceledException(
                    $"Loading cancelled. EstimatedProgress = {args.EstimatedProgress}; CanceledAt = {DateTime.Now}");
        }
    }

    // Callback for document saving. Throws OperationCanceledException only when enabled.
    public class SavingProgressCallback : IDocumentSavingCallback
    {
        private readonly DateTime _savingStartedAt = DateTime.Now;
        private readonly double _maxDurationSeconds;
        private readonly CancellationConfig _config;

        public SavingProgressCallback(double maxDurationSeconds, CancellationConfig config)
        {
            _maxDurationSeconds = maxDurationSeconds;
            _config = config;
        }

        public void Notify(DocumentSavingArgs args)
        {
            // If cancellation is disabled, simply return.
            if (!_config.EnableSavingCancellation)
                return;

            double elapsedSeconds = (DateTime.Now - _savingStartedAt).TotalSeconds;
            if (elapsedSeconds > _maxDurationSeconds)
                throw new OperationCanceledException(
                    $"Saving cancelled. EstimatedProgress = {args.EstimatedProgress}; CanceledAt = {DateTime.Now}");
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Path to the source document.
            string sourcePath = @"C:\Docs\BigDocument.docx";
            // Path to the output document.
            string outputPath = @"C:\Docs\BigDocument_Cancelled.docx";

            // Create configuration and enable cancellation for both stages.
            var cancellationConfig = new CancellationConfig
            {
                EnableLoadingCancellation = true,
                EnableSavingCancellation = true
            };

            // Set up load options with the custom loading callback.
            var loadOptions = new LoadOptions
            {
                ProgressCallback = new LoadingProgressCallback(maxDurationSeconds: 0.5, config: cancellationConfig)
            };

            Document doc = null;
            try
            {
                // Load the document using the configured load options.
                doc = new Document(sourcePath, loadOptions);
            }
            catch (OperationCanceledException ex)
            {
                Console.WriteLine($"Loading aborted: {ex.Message}");
                // Handle loading cancellation (e.g., cleanup, retry, etc.).
                return;
            }

            // Set up save options with the custom saving callback.
            var saveOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ProgressCallback = new SavingProgressCallback(maxDurationSeconds: 0.2, config: cancellationConfig)
            };

            try
            {
                // Save the document using the configured save options.
                doc.Save(outputPath, saveOptions);
            }
            catch (OperationCanceledException ex)
            {
                Console.WriteLine($"Saving aborted: {ex.Message}");
                // Handle saving cancellation (e.g., delete partial file, notify user, etc.).
            }
        }
    }
}
