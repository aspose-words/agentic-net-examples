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
        private int _partCount;

        public SavedDocumentPartRename(string baseFileName, DocumentSplitCriteria splitCriteria)
        {
            _baseFileName = baseFileName;
            _splitCriteria = splitCriteria;
            _partCount = 0;
        }

        // This method is called for each part that Aspose.Words is about to save.
        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // Determine a readable description of the split criteria.
            string partType = _splitCriteria switch
            {
                DocumentSplitCriteria.PageBreak => "Page",
                DocumentSplitCriteria.ColumnBreak => "Column",
                DocumentSplitCriteria.SectionBreak => "Section",
                DocumentSplitCriteria.HeadingParagraph => "HeadingParagraph",
                _ => "Part"
            };

            // Build a unique filename for the part.
            string partFileName = $"{_baseFileName}_part{++_partCount}_{partType}{Path.GetExtension(args.DocumentPartFileName)}";

            // Set the filename – Aspose.Words will write the part to this file.
            args.DocumentPartFileName = partFileName;

            // Alternatively you could provide a custom stream:
            // args.DocumentPartStream = new FileStream(Path.Combine(ArtifactsDir, partFileName), FileMode.Create);
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the source DOCX document.
            string inputPath = @"C:\Docs\SourceDocument.docx";

            // Base name for the split output files (without extension).
            string outputBaseName = "SplitDocument";

            // Load the DOCX document.
            Document doc = new Document(inputPath);

            // Configure HTML save options to split the document by section.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak,
                // Assign the custom callback to control part filenames.
                DocumentPartSavingCallback = new SavedDocumentPartRename(outputBaseName, DocumentSplitCriteria.SectionBreak)
            };

            // Save the document; Aspose.Words will create multiple HTML files according to the split criteria.
            doc.Save($"{outputBaseName}.html", saveOptions);
        }
    }
}
