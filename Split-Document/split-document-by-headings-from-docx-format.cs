using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOCX document.
        Document doc = new Document("Input.docx");

        // Configure HTML save options to split the document at heading paragraphs.
        HtmlSaveOptions options = new HtmlSaveOptions
        {
            // Split at headings (Heading 1, Heading 2, etc.).
            DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,
            // Define up to which heading level the split should occur.
            DocumentSplitHeadingLevel = 2,
            // Provide a callback to control the filenames of the generated parts.
            DocumentPartSavingCallback = new SavedDocumentPartRename("Output.html", DocumentSplitCriteria.HeadingParagraph)
        };

        // Save the document; Aspose.Words will create separate HTML files for each heading part.
        doc.Save("Output.html", options);
    }

    // Callback that customizes the filenames of each split document part.
    private class SavedDocumentPartRename : IDocumentPartSavingCallback
    {
        private readonly string _baseFileName;
        private readonly DocumentSplitCriteria _criteria;
        private int _partCounter;

        public SavedDocumentPartRename(string baseFileName, DocumentSplitCriteria criteria)
        {
            _baseFileName = baseFileName;
            _criteria = criteria;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // Determine a readable part type name.
            string partType = _criteria switch
            {
                DocumentSplitCriteria.PageBreak => "Page",
                DocumentSplitCriteria.ColumnBreak => "Column",
                DocumentSplitCriteria.SectionBreak => "Section",
                DocumentSplitCriteria.HeadingParagraph => "Heading",
                _ => "Part"
            };

            // Build a unique filename for the part.
            string partFileName = $"{Path.GetFileNameWithoutExtension(_baseFileName)}_part{++_partCounter}_{partType}{Path.GetExtension(args.DocumentPartFileName)}";

            // Set the filename (or you could set a custom stream instead).
            args.DocumentPartFileName = partFileName;
        }
    }
}
