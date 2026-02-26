using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class SplitDotx
{
    static void Main()
    {
        // Load the DOTX template file.
        Document doc = new Document("Template.dotx");

        // Configure HTML save options to split the document at each section break.
        HtmlSaveOptions options = new HtmlSaveOptions();
        options.DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak;

        // Assign a callback that will rename each generated part file.
        options.DocumentPartSavingCallback = new PartRenamer("SplitPart", options.DocumentSplitCriteria);

        // Save the document; Aspose.Words will create multiple HTML files according to the split criteria.
        doc.Save("Output.html", options);
    }

    // Callback implementation that customizes the file name of each document part.
    private class PartRenamer : IDocumentPartSavingCallback
    {
        private readonly string _baseName;
        private readonly DocumentSplitCriteria _criteria;
        private int _counter;

        public PartRenamer(string baseName, DocumentSplitCriteria criteria)
        {
            _baseName = baseName;
            _criteria = criteria;
        }

        public void DocumentPartSaving(DocumentPartSavingArgs args)
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

            // Build a new file name for the part.
            string newFileName = $"{_baseName}_Part{++_counter}_{partType}{Path.GetExtension(args.DocumentPartFileName)}";

            // Apply the new file name. Let Aspose handle the stream automatically.
            args.DocumentPartFileName = newFileName;
        }
    }
}
