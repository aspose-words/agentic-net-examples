using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Saving;

namespace DocxToTiffWatcher
{
    // Service that watches a folder and converts new DOCX files to TIFF.
    public class DocumentConversionService : IDisposable
    {
        private readonly string _watchFolder;
        private readonly string _outputFolder;
        private readonly FileSystemWatcher _watcher;
        private readonly SynchronizationContext _syncContext;
        private bool _disposed;

        public DocumentConversionService(string watchFolder, string outputFolder)
        {
            _watchFolder = watchFolder;
            _outputFolder = outputFolder;

            // Ensure output folder exists.
            Directory.CreateDirectory(_outputFolder);

            // FileSystemWatcher setup.
            _watcher = new FileSystemWatcher(_watchFolder, "*.docx")
            {
                NotifyFilter = NotifyFilters.FileName | NotifyFilters.CreationTime,
                EnableRaisingEvents = false,
                IncludeSubdirectories = false
            };
            _watcher.Created += OnCreated;

            // Use the thread pool context for async processing.
            _syncContext = new SynchronizationContext();
        }

        // Starts monitoring.
        public void Start()
        {
            _watcher.EnableRaisingEvents = true;
        }

        // Stops monitoring.
        public void Stop()
        {
            _watcher.EnableRaisingEvents = false;
        }

        // Handles new file creation.
        private void OnCreated(object sender, FileSystemEventArgs e)
        {
            // Queue the conversion to avoid blocking the watcher thread.
            ThreadPool.QueueUserWorkItem(_ => ProcessFile(e.FullPath));
        }

        // Performs the conversion from DOCX to TIFF.
        private void ProcessFile(string sourcePath)
        {
            // Wait until the file is accessible (some apps lock the file briefly).
            const int maxAttempts = 10;
            int attempts = 0;
            while (attempts < maxAttempts)
            {
                try
                {
                    // Attempt to open the file with read sharing.
                    using (FileStream stream = new FileStream(sourcePath, FileMode.Open, FileAccess.Read, FileShare.Read))
                    {
                        // Load the document using Aspose.Words constructor.
                        Document doc = new Document(stream);

                        // Build output file name with same base name but .tiff extension.
                        string fileName = Path.GetFileNameWithoutExtension(sourcePath);
                        string outputPath = Path.Combine(_outputFolder, fileName + ".tiff");

                        // Save the document as a multi‑page TIFF.
                        doc.Save(outputPath, SaveFormat.Tiff);
                    }
                    break; // Success, exit retry loop.
                }
                catch (IOException)
                {
                    // File might still be locked; wait and retry.
                    attempts++;
                    Thread.Sleep(500);
                }
                catch (Exception ex)
                {
                    // Log unexpected errors and abort processing for this file.
                    Console.Error.WriteLine($"Error converting '{sourcePath}': {ex.Message}");
                    break;
                }
            }

            if (attempts == maxAttempts)
            {
                Console.Error.WriteLine($"Failed to access '{sourcePath}' after multiple attempts.");
            }
        }

        // Dispose pattern to clean up resources.
        public void Dispose()
        {
            if (_disposed) return;
            _watcher.Dispose();
            _disposed = true;
        }
    }

    // Simple console host.
    internal static class Program
    {
        private static void Main(string[] args)
        {
            // Expect two arguments: folder to watch and folder for TIFF output.
            if (args.Length != 2)
            {
                Console.WriteLine("Usage: DocxToTiffWatcher <watchFolder> <outputFolder>");
                return;
            }

            string watchFolder = args[0];
            string outputFolder = args[1];

            using (var service = new DocumentConversionService(watchFolder, outputFolder))
            {
                service.Start();
                Console.WriteLine($"Watching '{watchFolder}' for new DOCX files. Press Enter to stop.");
                Console.ReadLine();
                service.Stop();
            }
        }
    }
}
