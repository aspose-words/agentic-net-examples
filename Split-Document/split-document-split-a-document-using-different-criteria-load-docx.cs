using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace DocumentSplitExample
{
    // Custom callback to rename each split part of the document.
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
            // Determine a readable part type based on the split criteria used.
            string partType = _splitCriteria switch
            {
                DocumentSplitCriteria.PageBreak => "Page",
                DocumentSplitCriteria.ColumnBreak => "Column",
                DocumentSplitCriteria.SectionBreak => "Section",
                DocumentSplitCriteria.HeadingParagraph => "Heading",
                _ => "Part"
            };

            // Build a new file name for the part.
            string newFileName = $"{_baseFileName}_part{++_partCounter}_{partType}{Path.GetExtension(args.DocumentPartFileName)}";

            // Assign the new file name.
            args.DocumentPartFileName = newFileName;

            // Alternatively, you could provide a custom stream:
            // args.DocumentPartStream = new FileStream(Path.Combine(ArtifactsDir, newFileName), FileMode.Create);
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the source DOCX document.
            string inputPath = @"C:\Docs\SourceDocument.docx";

            // Load the document from file.
            Document doc = new Document(inputPath);

            // Define the output base file name (without extension).
            string outputBaseName = "SplitDocument";

            // Configure HTML save options with split criteria.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                // Split the document at each section break.
                DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak,

                // Use the custom callback to control part file names.
                DocumentPartSavingCallback = new SavedDocumentPartRename(outputBaseName, DocumentSplitCriteria.SectionBreak)
            };

            // Save the document; Aspose.Words will create multiple HTML files according to the split criteria.
            string outputPath = Path.Combine(@"C:\Docs\Output", $"{outputBaseName}.html");
            doc.Save(outputPath, saveOptions);
        }
    }
}
