using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace DocumentSplitExample
{
    // Implements a callback to customize the filenames of the split parts.
    class SavedDocumentPartRename : IDocumentPartSavingCallback
    {
        private readonly string _baseFileName;
        private readonly DocumentSplitCriteria _criteria;
        private int _partIndex = 0;

        public SavedDocumentPartRename(string baseFileName, DocumentSplitCriteria criteria)
        {
            _baseFileName = baseFileName;
            _criteria = criteria;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // Determine a readable part type name.
            string partType = _criteria switch
            {
                DocumentSplitCriteria.PageBreak => "Page",
                DocumentSplitCriteria.ColumnBreak => "Column",
                DocumentSplitCriteria.SectionBreak => "Section",
                DocumentSplitCriteria.HeadingParagraph => "Heading",
                _ => "Part"
            };

            // Build a new filename for the part.
            string newFileName = $"{_baseFileName}_part{++_partIndex}_{partType}{Path.GetExtension(args.DocumentPartFileName)}";

            // Assign the new filename.
            args.DocumentPartFileName = newFileName;

            // Optionally, write to a custom stream (here we just let Aspose handle the file).
            // args.DocumentPartStream = new FileStream(Path.Combine(Path.GetDirectoryName(_baseFileName) ?? "", newFileName), FileMode.Create);
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the source DOCX file.
            string inputPath = @"C:\Docs\SourceDocument.docx";

            // Load the document from disk.
            Document doc = new Document(inputPath);

            // Configure HTML save options to split the document by section breaks.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak,
                // Optional: limit heading level when using HeadingParagraph criteria.
                DocumentSplitHeadingLevel = 2
            };

            // Attach the custom callback to control part filenames.
            string outputBase = @"C:\Docs\SplitOutput\Document";
            saveOptions.DocumentPartSavingCallback = new SavedDocumentPartRename(outputBase, saveOptions.DocumentSplitCriteria);

            // Ensure the output directory exists.
            Directory.CreateDirectory(Path.GetDirectoryName(outputBase) ?? string.Empty);

            // Save the document; Aspose will generate multiple HTML files according to the split criteria.
            doc.Save($"{outputBase}.html", saveOptions);
        }
    }
}
