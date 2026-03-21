using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsSplitExample
{
    // Implements the callback that controls how each document part is saved.
    internal class SavedDocumentPartRename : IDocumentPartSavingCallback
    {
        private readonly string _outputFolder;   // Folder where parts will be stored.
        private readonly string _baseFileName;   // Base name of the original document (without extension).
        private int _partCounter;                // Counter for generated part files.

        public SavedDocumentPartRename(string outputFolder, string baseFileName)
        {
            _outputFolder = outputFolder;
            _baseFileName = baseFileName;
            _partCounter = 0;
        }

        // Called by Aspose.Words for each part that is about to be saved.
        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // Build a unique file name for the part (e.g., MyDoc_part1.html).
            string partFileName = $"{_baseFileName}_part{++_partCounter}{Path.GetExtension(args.DocumentPartFileName)}";

            // Set the file name (without path) – Aspose.Words uses this for internal references.
            args.DocumentPartFileName = partFileName;

            // Create a full path on the folder and assign a stream for the part.
            string fullPath = Path.Combine(_outputFolder, partFileName);
            args.DocumentPartStream = new FileStream(fullPath, FileMode.Create, FileAccess.Write);

            // Ensure Aspose.Words closes the stream after writing.
            args.KeepDocumentPartStreamOpen = false;
        }
    }

    public class DocumentSplitter
    {
        /// <summary>
        /// Splits the input document by section breaks and saves each part to the specified folder.
        /// </summary>
        /// <param name="sourcePath">Full path to the source .docx (or any supported) file.</param>
        /// <param name="outputFolder">Folder where parts will be stored.</param>
        public static void SplitAndSaveBySections(string sourcePath, string outputFolder)
        {
            if (string.IsNullOrEmpty(sourcePath))
                throw new ArgumentException("Source path must be provided.", nameof(sourcePath));
            if (string.IsNullOrEmpty(outputFolder))
                throw new ArgumentException("Output folder must be provided.", nameof(outputFolder));

            // Ensure the output folder exists.
            Directory.CreateDirectory(outputFolder);

            // Load the document.
            Document doc = new Document(sourcePath);

            // Prepare HTML save options with section split criteria.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak
            };

            // Set the custom callback to control where each part is saved.
            string baseFileName = Path.GetFileNameWithoutExtension(sourcePath);
            saveOptions.DocumentPartSavingCallback = new SavedDocumentPartRename(outputFolder, baseFileName);

            // Save the main document.
            string mainOutputPath = Path.Combine(outputFolder, $"{baseFileName}.html");
            doc.Save(mainOutputPath, saveOptions);
        }

        // Example usage.
        public static void Main()
        {
            // Create a temporary source document with two sections.
            string tempFolder = Path.Combine(Path.GetTempPath(), "AsposeSplitDemo");
            Directory.CreateDirectory(tempFolder);
            string sourceDocPath = Path.Combine(tempFolder, "SampleDocument.docx");

            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Section 1 content.");
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.Writeln("Section 2 content.");
            doc.Save(sourceDocPath);

            // Output folder for split parts (local folder, not a network share).
            string outputFolder = Path.Combine(tempFolder, "WordParts");
            SplitAndSaveBySections(sourceDocPath, outputFolder);

            Console.WriteLine($"Document split and saved to: {outputFolder}");
        }
    }
}
