using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace DocumentSplitExample
{
    // Implements a callback to control how each split part is saved.
    public class SavedDocumentPartRename : IDocumentPartSavingCallback
    {
        private readonly string _baseFileName;
        private readonly DocumentSplitCriteria _splitCriteria;
        private int _partIndex = 0;

        public SavedDocumentPartRename(string baseFileName, DocumentSplitCriteria splitCriteria)
        {
            _baseFileName = baseFileName;
            _splitCriteria = splitCriteria;
        }

        // This method is called for each document part that Aspose.Words is about to save.
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

            // Build a new file name for the part.
            string newFileName = $"{Path.GetFileNameWithoutExtension(_baseFileName)}_part{++_partIndex}_{partType}{Path.GetExtension(args.DocumentPartFileName)}";

            // Set the new file name.
            args.DocumentPartFileName = newFileName;

            // Alternatively, provide a custom stream for the part.
            args.DocumentPartStream = new FileStream(Path.Combine(Path.GetDirectoryName(_baseFileName) ?? "", newFileName), FileMode.Create);
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the source DOCX document.
            string inputPath = @"C:\Docs\SourceDocument.docx";

            // Load the document using the Document constructor.
            Document doc = new Document(inputPath);

            // Prepare HTML save options with splitting at section breaks.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak,
                // Assign the custom callback to rename each split part.
                DocumentPartSavingCallback = new SavedDocumentPartRename(@"C:\Docs\SplitOutput.html", DocumentSplitCriteria.SectionBreak)
            };

            // Save the document; Aspose.Words will create multiple HTML files according to the split criteria.
            doc.Save(@"C:\Docs\SplitOutput.html", saveOptions);
        }
    }
}
