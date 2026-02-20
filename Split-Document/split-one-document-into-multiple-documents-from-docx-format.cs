using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOCX document.
        Document doc = new Document(@"Input\SourceDocument.docx");

        // Prepare HTML save options and configure splitting by section break.
        HtmlSaveOptions options = new HtmlSaveOptions
        {
            DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak,
            // The callback will control the filenames of the split parts.
            DocumentPartSavingCallback = new SavedDocumentPartRename("Output\\SplitDocument.html", DocumentSplitCriteria.SectionBreak)
        };

        // Save the document; Aspose.Words will create multiple HTML files according to the split criteria.
        doc.Save(@"Output\SplitDocument.html", options);
    }

    // Callback that assigns custom filenames to each document part created during the split operation.
    private class SavedDocumentPartRename : IDocumentPartSavingCallback
    {
        private readonly string _baseFileName;
        private readonly DocumentSplitCriteria _criteria;
        private int _partIndex;

        public SavedDocumentPartRename(string baseFileName, DocumentSplitCriteria criteria)
        {
            _baseFileName = baseFileName;
            _criteria = criteria;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // Determine a readable part type name based on the split criteria.
            string partType = _criteria switch
            {
                DocumentSplitCriteria.PageBreak => "Page",
                DocumentSplitCriteria.ColumnBreak => "Column",
                DocumentSplitCriteria.SectionBreak => "Section",
                DocumentSplitCriteria.HeadingParagraph => "Heading",
                _ => "Part"
            };

            // Build a unique filename for the current part.
            string partFileName = $"{Path.GetFileNameWithoutExtension(_baseFileName)}_part{++_partIndex}_{partType}{Path.GetExtension(args.DocumentPartFileName)}";

            // Set the filename (without path) for the part.
            args.DocumentPartFileName = partFileName;

            // Optionally, direct the part to a specific folder via a stream.
            string fullPath = Path.Combine(Path.GetDirectoryName(_baseFileName) ?? string.Empty, partFileName);
            args.DocumentPartStream = new FileStream(fullPath, FileMode.Create);
        }
    }
}
