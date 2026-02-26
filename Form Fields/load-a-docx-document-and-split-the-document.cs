using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class DocumentSplitter
{
    static void Main()
    {
        // Load the source DOCX document from the file system.
        Document doc = new Document("InputDocument.docx");

        // Configure HTML save options to split the document by section.
        HtmlSaveOptions options = new HtmlSaveOptions
        {
            DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak,
            // Assign a custom callback to control the naming of each split part.
            DocumentPartSavingCallback = new SavedDocumentPartRename("SplitDocument.html", DocumentSplitCriteria.SectionBreak)
        };

        // Save the document; Aspose.Words will create multiple HTML files according to the split criteria.
        doc.Save("SplitDocument.html", options);
    }

    // Callback that renames each split part generated during the save operation.
    private class SavedDocumentPartRename : IDocumentPartSavingCallback
    {
        private readonly string _baseFileName;
        private readonly DocumentSplitCriteria _splitCriteria;
        private int _partIndex = 0;

        public SavedDocumentPartRename(string baseFileName, DocumentSplitCriteria splitCriteria)
        {
            _baseFileName = baseFileName;
            _splitCriteria = splitCriteria;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // Determine a readable part type name based on the split criteria.
            string partType = _splitCriteria switch
            {
                DocumentSplitCriteria.PageBreak => "Page",
                DocumentSplitCriteria.ColumnBreak => "Column",
                DocumentSplitCriteria.SectionBreak => "Section",
                DocumentSplitCriteria.HeadingParagraph => "HeadingParagraph",
                _ => "Part"
            };

            // Build a unique filename for the current part.
            string partFileName = $"{Path.GetFileNameWithoutExtension(_baseFileName)}_Part{++_partIndex}_{partType}{Path.GetExtension(_baseFileName)}";

            // Set the filename for the part; Aspose.Words will write to this file.
            args.DocumentPartFileName = partFileName;

            // Optionally, you could provide a custom stream instead:
            // args.DocumentPartStream = new FileStream(partFileName, FileMode.Create);
        }
    }
}
