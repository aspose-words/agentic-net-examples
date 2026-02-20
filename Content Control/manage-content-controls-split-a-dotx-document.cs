using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace SplitDotxExample
{
    // Custom callback to control how each document part is saved when the document is split.
    public class PartRenamer : IDocumentPartSavingCallback
    {
        private readonly string _baseFileName;
        private readonly DocumentSplitCriteria _criteria;
        private int _partIndex = 0;

        public PartRenamer(string baseFileName, DocumentSplitCriteria criteria)
        {
            _baseFileName = baseFileName;
            _criteria = criteria;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // Determine a readable part type name based on the split criteria.
            string partType = _criteria switch
            {
                DocumentSplitCriteria.PageBreak => "Page",
                DocumentSplitCriteria.ColumnBreak => "Column",
                DocumentSplitCriteria.SectionBreak => "Section",
                DocumentSplitCriteria.HeadingParagraph => "Heading",
                _ => "Part"
            };

            // Build a new file name for the part.
            string newFileName = $"{Path.GetFileNameWithoutExtension(_baseFileName)}_{++_partIndex}_{partType}{Path.GetExtension(args.DocumentPartFileName)}";

            // Assign the new file name.
            args.DocumentPartFileName = newFileName;

            // Optionally, you could provide a custom stream instead of a file name:
            // args.DocumentPartStream = new FileStream(Path.Combine(Path.GetDirectoryName(_baseFileName) ?? "", newFileName), FileMode.Create);
        }
    }

    class Program
    {
        static void Main()
        {
            // Load the DOTX template.
            Document doc = new Document("Template.dotx");

            // Prepare HTML save options with split criteria.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                // Split the document into separate HTML files at each section break.
                DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak,

                // Use the custom callback to rename each part.
                DocumentPartSavingCallback = new PartRenamer("SplitResult.html", DocumentSplitCriteria.SectionBreak)
            };

            // Save the document; Aspose.Words will create multiple HTML files according to the split criteria.
            doc.Save("SplitResult.html", saveOptions);
        }
    }
}
