using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsSplitExample
{
    // Implements a callback to rename each split part when saving.
    public class SavedDocumentPartRename : IDocumentPartSavingCallback
    {
        private readonly string _baseFileName;
        private readonly DocumentSplitCriteria _splitCriteria;
        private int _partCounter = 0;

        public SavedDocumentPartRename(string baseFileName, DocumentSplitCriteria splitCriteria)
        {
            _baseFileName = baseFileName;
            _splitCriteria = splitCriteria;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // Determine a readable part type name.
            string partType = _splitCriteria switch
            {
                DocumentSplitCriteria.PageBreak => "Page",
                DocumentSplitCriteria.ColumnBreak => "Column",
                DocumentSplitCriteria.SectionBreak => "Section",
                DocumentSplitCriteria.HeadingParagraph => "HeadingParagraph",
                _ => "Part"
            };

            // Build a unique filename for the part.
            string partFileName = $"{Path.GetFileNameWithoutExtension(_baseFileName)}_part{++_partCounter}_{partType}{Path.GetExtension(args.DocumentPartFileName)}";

            // Set the filename where Aspose.Words will write this part.
            args.DocumentPartFileName = partFileName;

            // Optionally, write to a custom stream (here we use the default file handling).
            // args.DocumentPartStream = new FileStream(Path.Combine(Path.GetDirectoryName(_baseFileName) ?? "", partFileName), FileMode.Create);
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the source DOCX file.
            string sourcePath = @"C:\Docs\SourceDocument.docx";

            // Load the document from file.
            Document doc = new Document(sourcePath);

            // Prepare HTML save options with split criteria.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                // Split the document into separate files, one per section.
                DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak,

                // Assign the custom callback to control part filenames.
                DocumentPartSavingCallback = new SavedDocumentPartRename(@"C:\Docs\SplitOutput.html", DocumentSplitCriteria.SectionBreak)
            };

            // Save the document; Aspose.Words will generate multiple HTML files according to the split criteria.
            string outputPath = @"C:\Docs\SplitOutput.html";
            doc.Save(outputPath, saveOptions);
        }
    }
}
