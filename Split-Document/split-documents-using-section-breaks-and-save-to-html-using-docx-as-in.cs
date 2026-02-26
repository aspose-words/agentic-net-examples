using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace DocumentSplitExample
{
    // Custom callback to rename each HTML part generated when splitting by section breaks.
    class SavedDocumentPartRename : IDocumentPartSavingCallback
    {
        private readonly string _baseFileName;
        private int _partIndex = 0;

        public SavedDocumentPartRename(string baseFileName)
        {
            _baseFileName = baseFileName;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // Generate a sequential file name for each part.
            string partFileName = $"{Path.GetFileNameWithoutExtension(_baseFileName)}_Part{++_partIndex}{Path.GetExtension(args.DocumentPartFileName)}";

            // Set the file name for the part.
            args.DocumentPartFileName = partFileName;

            // Optionally, you could provide a custom stream instead of a file name:
            // args.DocumentPartStream = new FileStream(Path.Combine(Path.GetDirectoryName(_baseFileName), partFileName), FileMode.Create);
        }
    }

    class Program
    {
        static void Main()
        {
            // Input DOCX file path.
            string inputPath = @"C:\Docs\InputDocument.docx";

            // Output HTML base file name (the first part will use this name, additional parts will be renamed by the callback).
            string outputPath = @"C:\Docs\OutputDocument.html";

            // Load the DOCX document.
            Document doc = new Document(inputPath);

            // Configure HTML save options to split the document at each section break.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak,
                // Assign the custom callback to control part file naming.
                DocumentPartSavingCallback = new SavedDocumentPartRename(outputPath)
            };

            // Save the document; Aspose.Words will create multiple HTML files, one per section.
            doc.Save(outputPath, saveOptions);
        }
    }
}
