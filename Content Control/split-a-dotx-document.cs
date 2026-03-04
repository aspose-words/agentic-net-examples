using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class SplitDotx
{
    static void Main()
    {
        // Load the DOTX template from disk.
        Document doc = new Document("Template.dotx");

        // Configure HTML save options to split the document at each section break.
        HtmlSaveOptions options = new HtmlSaveOptions
        {
            DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak
        };

        // Set a custom callback that will rename each generated HTML part.
        string baseFileName = "SplitDocument.html";
        options.DocumentPartSavingCallback = new PartRenamer(baseFileName, options.DocumentSplitCriteria);

        // Save the document. Aspose.Words will create multiple HTML files according to the split criteria.
        doc.Save(Path.Combine("Output", baseFileName), options);
    }

    // Implements IDocumentPartSavingCallback to control the naming and storage of each document part.
    private class PartRenamer : IDocumentPartSavingCallback
    {
        private readonly string _baseName;
        private readonly DocumentSplitCriteria _criteria;
        private int _count;

        public PartRenamer(string baseName, DocumentSplitCriteria criteria)
        {
            _baseName = baseName;
            _criteria = criteria;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // Determine a readable part type based on the split criteria.
            string partType = _criteria switch
            {
                DocumentSplitCriteria.PageBreak => "Page",
                DocumentSplitCriteria.ColumnBreak => "Column",
                DocumentSplitCriteria.SectionBreak => "Section",
                DocumentSplitCriteria.HeadingParagraph => "Heading",
                _ => "Part"
            };

            // Construct a new filename for the part (e.g., SplitDocument_part1_Section.html).
            string newFileName = $"{Path.GetFileNameWithoutExtension(_baseName)}_part{++_count}_{partType}{Path.GetExtension(args.DocumentPartFileName)}";

            // Assign the new filename to the saving arguments.
            args.DocumentPartFileName = newFileName;

            // Optionally direct the output to a custom stream.
            args.DocumentPartStream = new FileStream(Path.Combine("Output", newFileName), FileMode.Create);

            // Ensure Aspose.Words closes the stream after writing.
            args.KeepDocumentPartStreamOpen = false;
        }
    }
}
