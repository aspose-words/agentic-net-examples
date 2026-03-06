using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsSplitExample
{
    // Callback that renames each split part of the document.
    class SavedDocumentPartRename : IDocumentPartSavingCallback
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
            // Determine the type of split (section, page, etc.) for naming.
            string partType = _splitCriteria switch
            {
                DocumentSplitCriteria.PageBreak => "Page",
                DocumentSplitCriteria.ColumnBreak => "Column",
                DocumentSplitCriteria.SectionBreak => "Section",
                DocumentSplitCriteria.HeadingParagraph => "Heading",
                _ => "Part"
            };

            // Build a unique file name for the current part.
            string partFileName = $"{Path.GetFileNameWithoutExtension(_baseFileName)}_Part{++_partIndex}_{partType}{Path.GetExtension(args.DocumentPartFileName)}";

            // Set the file name where Aspose.Words will write this part.
            args.DocumentPartFileName = partFileName;

            // Optionally, write to a custom stream instead of a file.
            // Here we let Aspose.Words handle the file creation, so we only set the name.
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the source DOCX file.
            string inputPath = @"C:\Docs\SourceDocument.docx";

            // Base name for the output HTML files (the main file and its parts).
            string outputHtml = @"C:\Docs\SplitDocument.html";

            // Load the DOCX document.
            Document doc = new Document(inputPath);

            // Configure HTML save options to split the document at each section break.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak,
                // Optional: customize how each part is named.
                DocumentPartSavingCallback = new SavedDocumentPartRename(outputHtml, DocumentSplitCriteria.SectionBreak)
            };

            // Save the document; Aspose.Words will create multiple HTML files according to the split criteria.
            doc.Save(outputHtml, saveOptions);
        }
    }
}
