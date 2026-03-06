using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace DocumentSplitter
{
    // Custom callback to control the naming of each split part when saving.
    public class SavedDocumentPartRename : IDocumentPartSavingCallback
    {
        private readonly string _baseFileName;
        private readonly DocumentSplitCriteria _splitCriteria;
        private int _partCounter = 0;

        public SavedDocumentPartRename(string baseFileName, DocumentSplitCriteria splitCriteria)
        {
            _baseFileName = baseFileName;
            _splitCriteria = splitCriteria;
        }

        // This method is invoked for each document part that Aspose.Words is about to save.
        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // Determine a friendly description of the split type.
            string partType = _splitCriteria switch
            {
                DocumentSplitCriteria.PageBreak => "Page",
                DocumentSplitCriteria.ColumnBreak => "Column",
                DocumentSplitCriteria.SectionBreak => "Section",
                DocumentSplitCriteria.HeadingParagraph => "HeadingParagraph",
                _ => "Part"
            };

            // Build a unique file name for the part.
            string partFileName = $"{_baseFileName}_part{++_partCounter}_{partType}{Path.GetExtension(args.DocumentPartFileName)}";

            // Specify the file name (or you could provide a stream instead).
            args.DocumentPartFileName = partFileName;

            // Ensure the stream is closed after saving (default behavior).
            args.KeepDocumentPartStreamOpen = false;
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the source DOCX document.
            string sourcePath = @"C:\Docs\SourceDocument.docx";

            // Load the document from the file system.
            Document doc = new Document(sourcePath);

            // Configure HTML save options to split the document at each section break.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak,
                // Assign the custom callback to control part file naming.
                DocumentPartSavingCallback = new SavedDocumentPartRename("SplitDocument", DocumentSplitCriteria.SectionBreak)
            };

            // Destination path for the first part (additional parts will be created automatically).
            string destinationPath = @"C:\Docs\SplitDocument.html";

            // Save the document; Aspose.Words will create multiple HTML files based on the split criteria.
            doc.Save(destinationPath, saveOptions);
        }
    }
}
