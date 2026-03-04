using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the existing DOCX document from the file system.
        Document doc = new Document("Input.docx");

        // Configure HTML save options to split the document at each section break.
        HtmlSaveOptions options = new HtmlSaveOptions
        {
            DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak,
            DocumentPartSavingCallback = new PartRenamer("SplitDocument")
        };

        // Save the document. Aspose.Words will create multiple HTML files,
        // one for each section, using the callback to name them.
        doc.Save("SplitDocument.html", options);
    }

    // Callback that assigns custom filenames to each split part.
    private class PartRenamer : IDocumentPartSavingCallback
    {
        private readonly string _baseName;
        private int _counter;

        public PartRenamer(string baseName)
        {
            _baseName = baseName;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // Preserve the original extension (e.g., .html) and create a sequential name.
            string extension = Path.GetExtension(args.DocumentPartFileName);
            string newFileName = $"{_baseName}_part_{++_counter}{extension}";

            // Set the new filename for the part.
            args.DocumentPartFileName = newFileName;

            // Ensure the stream used for this part is closed after saving.
            args.KeepDocumentPartStreamOpen = false;
        }
    }
}
