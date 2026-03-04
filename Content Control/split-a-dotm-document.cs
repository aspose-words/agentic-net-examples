using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsSplitDotm
{
    // Demonstrates how to split a DOTM (macro‑enabled template) into separate HTML parts.
    class Program
    {
        static void Main()
        {
            // Path to the source DOTM file.
            string inputPath = @"Input.dotm";

            // Load the document. The constructor automatically detects the DOTM format.
            Document doc = new Document(inputPath);

            // Base name for the output HTML files.
            string outFileName = "SplitDocument.html";

            // Folder where the split parts will be written.
            string outputDir = @"Output";
            Directory.CreateDirectory(outputDir);

            // Configure HTML save options to split the document at each section break.
            HtmlSaveOptions options = new HtmlSaveOptions
            {
                DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak,
                // Assign a callback to control the naming of each split part.
                DocumentPartSavingCallback = new SavedDocumentPartRename(outFileName, DocumentSplitCriteria.SectionBreak)
            };

            // Save the document. Aspose.Words will invoke the callback for each part.
            doc.Save(Path.Combine(outputDir, outFileName), options);
        }
    }

    // Callback that sets custom file names (and streams) for each document part created during saving.
    internal class SavedDocumentPartRename : IDocumentPartSavingCallback
    {
        private readonly string _baseFileName;
        private readonly DocumentSplitCriteria _splitCriteria;
        private int _partCount;

        public SavedDocumentPartRename(string baseFileName, DocumentSplitCriteria splitCriteria)
        {
            _baseFileName = baseFileName;
            _splitCriteria = splitCriteria;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // Determine a readable description of the split type.
            string partType = _splitCriteria switch
            {
                DocumentSplitCriteria.PageBreak => "Page",
                DocumentSplitCriteria.ColumnBreak => "Column",
                DocumentSplitCriteria.SectionBreak => "Section",
                DocumentSplitCriteria.HeadingParagraph => "HeadingParagraph",
                _ => "Part"
            };

            // Build a unique file name for this part.
            string partFileName = $"{_baseFileName} part {++_partCount}, of type {partType}{Path.GetExtension(args.DocumentPartFileName)}";

            // Set the file name.
            args.DocumentPartFileName = partFileName;

            // Optionally, provide a custom stream (here we write directly to a file in the output folder).
            string fullPath = Path.Combine("Output", partFileName);
            args.DocumentPartStream = new FileStream(fullPath, FileMode.Create);

            // Keep the stream closed after Aspose.Words finishes writing.
            args.KeepDocumentPartStreamOpen = false;
        }
    }
}
