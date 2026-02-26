using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class SplitDocumentExample
{
    static void Main()
    {
        // Load the source DOCX file.
        Document doc = new Document("Input.docx");

        // Configure HTML save options to split the document at each section break.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak;

        // Assign a callback that will give each split part a custom file name.
        saveOptions.DocumentPartSavingCallback = new PartRenamer("SplitDocument");

        // Save the document. Aspose.Words will invoke the callback for each part.
        doc.Save("Output.html", saveOptions);
    }

    // Callback implementation that renames each document part.
    private class PartRenamer : IDocumentPartSavingCallback
    {
        private readonly string _baseName;
        private int _partIndex;

        public PartRenamer(string baseName)
        {
            _baseName = baseName;
            _partIndex = 0;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // Create a unique file name for the current part.
            string newFileName = $"{_baseName}_Part{++_partIndex}{Path.GetExtension(args.DocumentPartFileName)}";

            // Set the new file name. Aspose.Words will write the part to this file.
            args.DocumentPartFileName = newFileName;

            // If you prefer to write to a custom stream, uncomment the following lines:
            // args.DocumentPartStream = new FileStream(newFileName, FileMode.Create);
            // args.KeepDocumentPartStreamOpen = false;
        }
    }
}
