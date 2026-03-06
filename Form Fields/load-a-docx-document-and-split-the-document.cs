using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class DocumentSplitter
{
    static void Main()
    {
        // Load the source DOCX document.
        // Uses the Document(string) constructor as defined in the API.
        Document doc = new Document("InputDocument.docx");

        // Prepare HTML save options to split the document by section.
        HtmlSaveOptions options = new HtmlSaveOptions
        {
            DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak,
            // Assign a custom callback to control the naming of each split part.
            DocumentPartSavingCallback = new SavedDocumentPartRename("SplitOutput", DocumentSplitCriteria.SectionBreak)
        };

        // Save the document. The save operation will produce multiple HTML files,
        // one for each section, using the logic defined in SavedDocumentPartRename.
        doc.Save("SplitOutput.html", options);
    }

    // Callback class that renames each split part of the document.
    private class SavedDocumentPartRename : IDocumentPartSavingCallback
    {
        private readonly string _baseFileName;
        private readonly DocumentSplitCriteria _splitCriteria;
        private int _partIndex;

        public SavedDocumentPartRename(string baseFileName, DocumentSplitCriteria splitCriteria)
        {
            _baseFileName = baseFileName;
            _splitCriteria = splitCriteria;
            _partIndex = 0;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // Determine a readable part type name.
            string partType = _splitCriteria switch
            {
                DocumentSplitCriteria.PageBreak => "Page",
                DocumentSplitCriteria.ColumnBreak => "Column",
                DocumentSplitCriteria.SectionBreak => "Section",
                DocumentSplitCriteria.HeadingParagraph => "HeadingParagraph",
                _ => "Part"
            };

            // Build a unique filename for the current part.
            string partFileName = $"{_baseFileName}_Part{++_partIndex}_{partType}{Path.GetExtension(args.DocumentPartFileName)}";

            // Set the filename for the output part.
            args.DocumentPartFileName = partFileName;

            // Optionally, you could provide a custom stream instead:
            // args.DocumentPartStream = new FileStream(partFileName, FileMode.Create);
        }
    }
}
