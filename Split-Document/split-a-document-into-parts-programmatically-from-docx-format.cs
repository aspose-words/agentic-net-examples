using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOCX document.
        Document doc = new Document("Input.docx");

        // Configure HTML save options to split the document at each section break.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak;

        // Assign a callback that will rename each generated HTML part.
        saveOptions.DocumentPartSavingCallback = new PartRenamer("OutputPart", saveOptions.DocumentSplitCriteria);

        // Save the document. Aspose.Words will create multiple HTML files according to the split criteria.
        doc.Save("Output.html", saveOptions);
    }

    // Callback implementation that controls the filenames (and streams) of the split parts.
    private class PartRenamer : IDocumentPartSavingCallback
    {
        private readonly string _baseName;
        private readonly DocumentSplitCriteria _criteria;
        private int _counter;

        public PartRenamer(string baseName, DocumentSplitCriteria criteria)
        {
            _baseName = baseName;
            _criteria = criteria;
            _counter = 0;
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

            // Build a unique filename for the part (extension is preserved).
            string newFileName = $"{_baseName}_{++_counter}_{partType}{Path.GetExtension(args.DocumentPartFileName)}";

            // Set the new filename.
            args.DocumentPartFileName = newFileName;

            // Optionally provide a custom stream for the part.
            args.DocumentPartStream = new FileStream(newFileName, FileMode.Create);
        }
    }
}
