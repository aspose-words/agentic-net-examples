using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsSplitExample
{
    // Custom callback to control the filenames of the split document parts.
    class DocumentPartRenamer : IDocumentPartSavingCallback
    {
        private readonly string _baseFileName;
        private int _partIndex = 0;

        public DocumentPartRenamer(string baseFileName)
        {
            _baseFileName = baseFileName;
        }

        // This method is called for each part that Aspose.Words is about to save.
        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // Generate a new filename for the part, e.g. "MyDocument part 1.html".
            string partFileName = $"{_baseFileName} part {++_partIndex}{Path.GetExtension(args.DocumentPartFileName)}";

            // Set the filename; Aspose.Words will write the part to this file.
            args.DocumentPartFileName = partFileName;

            // Optionally, you could provide a custom stream instead of a filename:
            // args.DocumentPartStream = new FileStream(partFileName, FileMode.Create);
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the source DOCX file.
            string sourcePath = @"C:\Docs\SourceDocument.docx";

            // Load the DOCX document.
            Document doc = new Document(sourcePath);

            // Configure HTML save options to split the document at each section break.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak
            };

            // Assign the custom callback to rename each split part.
            saveOptions.DocumentPartSavingCallback = new DocumentPartRenamer("MyDocument");

            // Save the document. The main file name is used for the first part;
            // additional parts will be created according to the callback logic.
            string outputPath = @"C:\Docs\MyDocument.html";
            doc.Save(outputPath, saveOptions);

            Console.WriteLine("Document split and saved successfully.");
        }
    }
}
