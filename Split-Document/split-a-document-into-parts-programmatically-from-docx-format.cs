using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace DocumentSplitExample
{
    // Implements a callback to customize the filenames of the split document parts.
    public class SavedDocumentPartRename : IDocumentPartSavingCallback
    {
        private readonly string _baseFileName;
        private readonly DocumentSplitCriteria _splitCriteria;
        private int _partCounter;

        public SavedDocumentPartRename(string baseFileName, DocumentSplitCriteria splitCriteria)
        {
            _baseFileName = baseFileName;
            _splitCriteria = splitCriteria;
            _partCounter = 0;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // Determine a readable part type based on the split criteria.
            string partType = _splitCriteria switch
            {
                DocumentSplitCriteria.PageBreak => "Page",
                DocumentSplitCriteria.ColumnBreak => "Column",
                DocumentSplitCriteria.SectionBreak => "Section",
                DocumentSplitCriteria.HeadingParagraph => "Heading",
                _ => "Part"
            };

            // Build a new filename for the part.
            string newFileName = $"{_baseFileName}_part{++_partCounter}_{partType}{Path.GetExtension(args.DocumentPartFileName)}";

            // Option 1 – set the filename directly.
            args.DocumentPartFileName = newFileName;

            // Option 2 – provide a custom stream (optional, shown for completeness).
            // args.DocumentPartStream = new FileStream(Path.Combine(ArtifactsDir, newFileName), FileMode.Create);
        }
    }

    class Program
    {
        // Adjust these paths as needed.
        private const string InputDocxPath = @"C:\Docs\InputDocument.docx";
        private const string OutputHtmlBaseName = "SplitDocument";

        static void Main()
        {
            // Load the source DOCX document.
            Document doc = new Document(InputDocxPath);

            // Configure HTML save options to split the document at each section break.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak,
                DocumentPartSavingCallback = new SavedDocumentPartRename(OutputHtmlBaseName, DocumentSplitCriteria.SectionBreak)
            };

            // Save the document; Aspose.Words will create multiple HTML files according to the split criteria.
            doc.Save($"{OutputHtmlBaseName}.html", saveOptions);
        }
    }
}
