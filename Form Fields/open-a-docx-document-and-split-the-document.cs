using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace DocumentSplitExample
{
    // Implements the callback that controls how each split part is saved.
    class SavedDocumentPartRename : IDocumentPartSavingCallback
    {
        private readonly string _baseFileName;
        private readonly DocumentSplitCriteria _splitCriteria;
        private int _partCount;

        public SavedDocumentPartRename(string baseFileName, DocumentSplitCriteria splitCriteria)
        {
            _baseFileName = baseFileName;
            _splitCriteria = splitCriteria;
            _partCount = 0;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // Determine a readable part type name.
            string partType = _splitCriteria switch
            {
                DocumentSplitCriteria.PageBreak => "Page",
                DocumentSplitCriteria.ColumnBreak => "Column",
                DocumentSplitCriteria.SectionBreak => "Section",
                DocumentSplitCriteria.HeadingParagraph => "Heading",
                _ => "Part"
            };

            // Build a unique file name for the part.
            string partFileName = $"{Path.GetFileNameWithoutExtension(_baseFileName)}_part{++_partCount}_{partType}{Path.GetExtension(args.DocumentPartFileName)}";

            // Set the file name (Aspose.Words will write the part to this file).
            args.DocumentPartFileName = partFileName;

            // Alternatively you could provide a custom stream:
            // args.DocumentPartStream = new FileStream(partFileName, FileMode.Create);
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the source DOCX file.
            string inputPath = @"C:\Docs\SourceDocument.docx";

            // Base name for the split output files.
            string outputBase = @"C:\Docs\SplitDocument.html";

            // Load the document from file.
            Document doc = new Document(inputPath);

            // Configure HTML save options to split the document at each section break.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak,
                DocumentPartSavingCallback = new SavedDocumentPartRename(outputBase, DocumentSplitCriteria.SectionBreak)
            };

            // Save the document; Aspose.Words will create multiple HTML files according to the split criteria.
            doc.Save(outputBase, saveOptions);
        }
    }
}
