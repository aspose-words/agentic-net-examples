using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace DocumentSplitExample
{
    // Callback that assigns a custom name to each split part.
    public class PartNamingCallback : IDocumentPartSavingCallback
    {
        private readonly string _baseFileName;
        private int _partIndex = 0;

        public PartNamingCallback(string baseFileName)
        {
            _baseFileName = baseFileName;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // Increment part counter.
            _partIndex++;

            // Build a new file name for the part, preserving the original extension.
            string newFileName = $"{_baseFileName}_Part{_partIndex}{Path.GetExtension(args.DocumentPartFileName)}";

            // Assign the new file name.
            args.DocumentPartFileName = newFileName;

            // Optionally, you could provide a custom stream instead:
            // args.DocumentPartStream = new FileStream(Path.Combine(Path.GetDirectoryName(newFileName), newFileName), FileMode.Create);
        }
    }

    class Program
    {
        static void Main()
        {
            // Load the source DOCX document.
            Document doc = new Document("InputDocument.docx");

            // Configure HTML save options to split the document at each heading paragraph.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,
                // Assign the callback that will name each split part.
                DocumentPartSavingCallback = new PartNamingCallback("SplitDocument")
            };

            // Save the document. Aspose.Words will create multiple HTML files,
            // one for each heading paragraph, using the names supplied by the callback.
            doc.Save("SplitDocument.html", saveOptions);
        }
    }
}
