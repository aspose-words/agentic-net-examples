using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace SplitDocumentExample
{
    // Callback that saves the main combined file and each split part with deterministic names.
    class CustomDocumentPartSavingCallback : IDocumentPartSavingCallback
    {
        private readonly string _outputDir;
        private readonly string _mainFileName; // file name (with extension) for the combined document
        private int _partIndex = 0;

        public CustomDocumentPartSavingCallback(string outputDir, string mainFileName)
        {
            _outputDir = outputDir;
            _mainFileName = mainFileName;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // First part corresponds to the main file passed to Document.Save().
            if (_partIndex == 0)
            {
                // Keep the original main file name.
                args.DocumentPartFileName = _mainFileName;
                args.DocumentPartStream = new FileStream(Path.Combine(_outputDir, _mainFileName), FileMode.Create);
            }
            else
            {
                // Subsequent parts are saved as Page_1.html, Page_2.html, …
                string partFileName = $"Page_{_partIndex}.html";
                args.DocumentPartFileName = partFileName;
                args.DocumentPartStream = new FileStream(Path.Combine(_outputDir, partFileName), FileMode.Create);
            }

            // Close the stream after Aspose.Words finishes writing.
            args.KeepDocumentPartStreamOpen = false;
            _partIndex++;
        }
    }

    public class Program
    {
        // Directory where output files will be written.
        private static readonly string ArtifactsDir = Path.Combine(Environment.CurrentDirectory, "Artifacts");

        public static void Main()
        {
            // Ensure a clean output folder.
            if (Directory.Exists(ArtifactsDir))
                Directory.Delete(ArtifactsDir, true);
            Directory.CreateDirectory(ArtifactsDir);

            // Build a sample document with three pages.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Page 1 content.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Page 2 content.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Page 3 content.");

            // Name of the combined (first) HTML file.
            const string combinedFileName = "Combined.html";

            // Set up HTML save options to split the document at page breaks.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                DocumentSplitCriteria = DocumentSplitCriteria.PageBreak,
                DocumentPartSavingCallback = new CustomDocumentPartSavingCallback(ArtifactsDir, combinedFileName)
            };

            // Save the document; the callback will create the combined file and the split pages.
            string combinedFilePath = Path.Combine(ArtifactsDir, combinedFileName);
            doc.Save(combinedFilePath, saveOptions);

            // Verify that the expected number of split files were created (combined + one per page break).
            string[] pageFiles = Directory.GetFiles(ArtifactsDir, "Page_*.html");
            if (pageFiles.Length != doc.PageCount - 1) // page breaks = pages - 1
                throw new InvalidOperationException($"Expected {doc.PageCount - 1} page files, but found {pageFiles.Length}.");

            // Verify that the main combined file exists.
            if (!File.Exists(combinedFilePath))
                throw new FileNotFoundException("The combined HTML file was not created.", combinedFilePath);
        }
    }
}
