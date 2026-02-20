using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsSplitExample
{
    // Custom callback to control the filenames of each split part (chapter).
    class ChapterPartRenamer : IDocumentPartSavingCallback
    {
        private readonly string _baseFileName;
        private int _chapterIndex = 0;

        public ChapterPartRenamer(string baseFileName)
        {
            _baseFileName = baseFileName;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // Increment chapter counter and build a new filename.
            _chapterIndex++;
            string extension = Path.GetExtension(args.DocumentPartFileName);
            string newFileName = $"{_baseFileName}_Chapter{_chapterIndex}{extension}";

            // Set the new filename for the part.
            args.DocumentPartFileName = newFileName;

            // Optionally, you could provide a custom stream instead of a file name:
            // args.DocumentPartStream = new FileStream(Path.Combine(outputFolder, newFileName), FileMode.Create);
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the source DOCX file.
            string sourceDocxPath = @"C:\Docs\SourceDocument.docx";

            // Load the DOCX document.
            Document doc = new Document(sourceDocxPath);

            // Configure save options to split the document into chapters.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Epub)
            {
                // Split at heading paragraphs (e.g., Heading 1, Heading 2, etc.).
                DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,

                // Define up to which heading level the split should occur.
                DocumentSplitHeadingLevel = 2,

                // Optional: set a custom base filename for the generated parts.
                DocumentPartSavingCallback = new ChapterPartRenamer("MyBook")
            };

            // Output EPUB file (the main file that references the split parts).
            string outputEpubPath = @"C:\Docs\MyBook.epub";

            // Save the document; Aspose.Words will create separate EPUB parts for each chapter.
            doc.Save(outputEpubPath, saveOptions);
        }
    }
}
