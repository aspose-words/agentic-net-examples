using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace DocumentSplitExample
{
    // Custom callback to control the naming of each split part.
    public class SavedDocumentPartRename : IDocumentPartSavingCallback
    {
        private readonly string _baseFileName;
        private readonly DocumentSplitCriteria _criteria;
        private int _partCounter;

        public SavedDocumentPartRename(string baseFileName, DocumentSplitCriteria criteria)
        {
            _baseFileName = baseFileName;
            _criteria = criteria;
            _partCounter = 0;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // Determine a readable description of the split criteria.
            string partType = _criteria switch
            {
                DocumentSplitCriteria.PageBreak => "Page",
                DocumentSplitCriteria.ColumnBreak => "Column",
                DocumentSplitCriteria.SectionBreak => "Section",
                DocumentSplitCriteria.HeadingParagraph => "Heading",
                _ => "Part"
            };

            // Build a unique file name for the current part.
            string partFileName = $"{_baseFileName}_part{++_partCounter}_{partType}{Path.GetExtension(args.DocumentPartFileName)}";

            // Option 1: set the file name directly.
            args.DocumentPartFileName = partFileName;

            // Option 2: provide a custom stream (uncomment if you prefer streams).
            // args.DocumentPartStream = new FileStream(Path.Combine(Path.GetDirectoryName(_baseFileName) ?? "", partFileName), FileMode.Create);
        }
    }

    class Program
    {
        static void Main()
        {
            // Load the source DOCX document.
            Document doc = new Document("InputDocument.docx");

            // Prepare HTML save options with the desired split criteria.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                // Choose one or combine multiple criteria using bitwise OR.
                DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak |
                                        DocumentSplitCriteria.HeadingParagraph
            };

            // Assign the custom callback to control part file names.
            saveOptions.DocumentPartSavingCallback = new SavedDocumentPartRename("SplitDocument", saveOptions.DocumentSplitCriteria);

            // Save the document; Aspose.Words will automatically split it according to the criteria.
            doc.Save("SplitDocument.html", saveOptions);
        }
    }
}
