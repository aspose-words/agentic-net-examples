using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsSplitExample
{
    // Callback that controls how each document part is saved when the document is split.
    internal class SavedDocumentPartRename : IDocumentPartSavingCallback
    {
        private readonly string _outputDirectory;
        private readonly string _baseFileName;
        private readonly DocumentSplitCriteria _splitCriteria;
        private int _partCount;

        public SavedDocumentPartRename(string outputDirectory, string baseFileName, DocumentSplitCriteria splitCriteria)
        {
            _outputDirectory = outputDirectory;
            _baseFileName = baseFileName;
            _splitCriteria = splitCriteria;
            _partCount = 0;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // Determine a readable part type name (optional, used only for naming).
            string partType = _splitCriteria switch
            {
                DocumentSplitCriteria.PageBreak => "Page",
                DocumentSplitCriteria.ColumnBreak => "Column",
                DocumentSplitCriteria.SectionBreak => "Section",
                DocumentSplitCriteria.HeadingParagraph => "Heading",
                _ => "Part"
            };

            // Build a unique file name for the part.
            string partFileName = $"{_baseFileName}_part{++_partCount}_{partType}{Path.GetExtension(args.DocumentPartFileName)}";

            // Set the file name (used when saving to a file) and the stream (used for direct writing).
            args.DocumentPartFileName = partFileName;
            args.DocumentPartStream = new FileStream(Path.Combine(_outputDirectory, partFileName), FileMode.Create);
            args.KeepDocumentPartStreamOpen = false; // close after each part is written.
        }
    }

    public static class DocumentSplitter
    {
        // Splits each document in the inputFiles array according to the specified criteria
        // and stores the resulting parts in the corresponding outputDirectories.
        public static void SplitDocuments(string[] inputFiles, string[] outputDirectories)
        {
            if (inputFiles == null) throw new ArgumentNullException(nameof(inputFiles));
            if (outputDirectories == null) throw new ArgumentNullException(nameof(outputDirectories));
            if (inputFiles.Length != outputDirectories.Length)
                throw new ArgumentException("The number of input files must match the number of output directories.");

            for (int i = 0; i < inputFiles.Length; i++)
            {
                string inputPath = inputFiles[i];
                string outputDir = outputDirectories[i];

                // Ensure the output directory exists.
                Directory.CreateDirectory(outputDir);

                // Load the source document.
                Document doc = new Document(inputPath);

                // Prepare save options for HTML with splitting at section breaks.
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak
                };

                // Base file name for the main HTML file (without extension).
                string baseFileName = Path.GetFileNameWithoutExtension(inputPath) + ".html";

                // Assign the custom callback that will name and stream each part.
                saveOptions.DocumentPartSavingCallback = new SavedDocumentPartRename(
                    outputDir,
                    baseFileName,
                    saveOptions.DocumentSplitCriteria);

                // Save the document; Aspose.Words will invoke the callback for each part.
                string mainOutputPath = Path.Combine(outputDir, baseFileName);
                doc.Save(mainOutputPath, saveOptions);
            }
        }
    }

    // Example usage.
    class Program
    {
        static void Main()
        {
            // Define source documents.
            string[] sources = new[]
            {
                @"C:\Docs\Report1.docx",
                @"C:\Docs\Report2.docx"
            };

            // Define where each document's split parts should be stored.
            string[] destinations = new[]
            {
                @"C:\SplitOutput\Report1",
                @"C:\SplitOutput\Report2"
            };

            // Perform the split operation.
            DocumentSplitter.SplitDocuments(sources, destinations);
        }
    }
}
