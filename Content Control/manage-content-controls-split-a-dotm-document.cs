using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace SplitDotmExample
{
    // Callback that renames each split part when saving.
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
            // Determine a readable part type name.
            string partType = _splitCriteria switch
            {
                DocumentSplitCriteria.PageBreak => "Page",
                DocumentSplitCriteria.ColumnBreak => "Column",
                DocumentSplitCriteria.SectionBreak => "Section",
                DocumentSplitCriteria.HeadingParagraph => "Heading",
                _ => "Part"
            };

            // Build a new file name for the part.
            string newFileName = $"{Path.GetFileNameWithoutExtension(_baseFileName)}_part{++_partCounter}_{partType}{Path.GetExtension(args.DocumentPartFileName)}";

            // Set the file name – Aspose.Words will write the part to this file.
            args.DocumentPartFileName = newFileName;

            // Alternatively you could provide a custom stream:
            // args.DocumentPartStream = new FileStream(Path.Combine(Path.GetDirectoryName(_baseFileName) ?? "", newFileName), FileMode.Create);
        }
    }

    class Program
    {
        static void Main()
        {
            // Load the DOTM template. The format is detected automatically.
            Document doc = new Document("Template.dotm");

            // Prepare HTML save options with split criteria.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                // Split the document into separate HTML files at each section break.
                DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak,

                // Use the custom callback to control part file names.
                DocumentPartSavingCallback = new SavedDocumentPartRename("Output.html", DocumentSplitCriteria.SectionBreak)
            };

            // Save the document. Each section will be saved as a separate HTML file.
            doc.Save("Output.html", saveOptions);
        }
    }
}
