using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsSplitExample
{
    // Implements the callback that controls how each document part is saved.
    internal class SavedDocumentPartRename : IDocumentPartSavingCallback
    {
        private readonly string _outputFolder;   // Network share folder where parts will be stored.
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

            // Create a full path on the network share and assign a stream for the part.
            string fullPath = Path.Combine(_outputFolder, partFileName);
            args.DocumentPartStream = new FileStream(fullPath, FileMode.Create, FileAccess.Write);

            // Ensure Aspose.Words closes the stream after writing.
            args.KeepDocumentPartStreamOpen = false;
        }
    }

    public class DocumentSplitter
    {
        /// <summary>
        /// Splits the input document by section breaks and saves each part to the specified network share.
        /// </summary>
        /// <param name="sourcePath">Full path to the source .docx (or any supported) file.</param>
        /// <param name="networkShareFolder">UNC path to the network share folder (e.g., \\Server\Share\Folder).</param>
        public static void SplitAndSaveBySections(string sourcePath, string networkShareFolder)
        {
            // Validate inputs.
            if (string.IsNullOrEmpty(sourcePath))
                throw new ArgumentException("Source path must be provided.", nameof(sourcePath));
            if (string.IsNullOrEmpty(networkShareFolder))
                throw new ArgumentException("Network share folder must be provided.", nameof(networkShareFolder));

            // Ensure the network folder exists; create it if necessary.
            if (!Directory.Exists(networkShareFolder))
                Directory.CreateDirectory(networkShareFolder);

            // Load the document (lifecycle rule: load).
            Document doc = new Document(sourcePath);

            // Prepare HTML save options with section split criteria.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak
            };

            // Set the custom callback to control where each part is saved.
            string baseFileName = Path.GetFileNameWithoutExtension(sourcePath);
            saveOptions.DocumentPartSavingCallback = new SavedDocumentPartRename(networkShareFolder, baseFileName);

            // Save the document (lifecycle rule: save). The main file is also written to the network share.
            string mainOutputPath = Path.Combine(networkShareFolder, $"{baseFileName}.html");
            doc.Save(mainOutputPath, saveOptions);
        }

        // Example usage.
        public static void Main()
        {
            // Local source document.
            string localDocPath = @"C:\Docs\SampleDocument.docx";

            // UNC network share where parts will be stored.
            string networkShare = @"\\MyServer\SharedDocs\WordParts";

            SplitAndSaveBySections(localDocPath, networkShare);

            Console.WriteLine("Document split and saved to network share successfully.");
        }
    }
}
