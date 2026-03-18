using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace DocumentPartSavingDemo
{
    class Program
    {
        static void Main()
        {
            // Load the source document.
            Document doc = new Document(@"MyDir\Rendering.docx");

            // Base name for the output HTML file.
            string baseOutputFileName = "DocumentParts.html";

            // Configure HTML save options.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                // Split the document into separate HTML files per section.
                DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak,

                // Assign a custom callback to rename each document part.
                DocumentPartSavingCallback = new SavedDocumentPartRename(baseOutputFileName, DocumentSplitCriteria.SectionBreak)
            };

            // Save the document using the configured options.
            doc.Save(@"ArtifactsDir\" + baseOutputFileName, saveOptions);
        }
    }

    /// <summary>
    /// Custom callback that renames each document part generated during HTML export.
    /// </summary>
    internal class SavedDocumentPartRename : IDocumentPartSavingCallback
    {
        private readonly string _baseFileName;
        private readonly DocumentSplitCriteria _splitCriteria;
        private int _partCounter;

        public SavedDocumentPartRename(string baseFileName, DocumentSplitCriteria splitCriteria)
        {
            _baseFileName = baseFileName;
            _splitCriteria = splitCriteria;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // Determine the type of part based on the split criteria.
            string partType = _splitCriteria switch
            {
                DocumentSplitCriteria.PageBreak => "Page",
                DocumentSplitCriteria.ColumnBreak => "Column",
                DocumentSplitCriteria.SectionBreak => "Section",
                DocumentSplitCriteria.HeadingParagraph => "Heading",
                _ => "Part"
            };

            // Build a unique file name for the part.
            string newFileName = $"{Path.GetFileNameWithoutExtension(_baseFileName)}_Part{++_partCounter}_{partType}{Path.GetExtension(args.DocumentPartFileName)}";

            // Set the new file name. Aspose.Words will use this name when saving the part.
            args.DocumentPartFileName = newFileName;

            // Optionally, you could provide a custom stream instead of a file name:
            // args.DocumentPartStream = new FileStream(Path.Combine("ArtifactsDir", newFileName), FileMode.Create);
        }
    }
}
