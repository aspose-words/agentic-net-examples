using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace ChapterSplitExample
{
    class Program
    {
        static void Main()
        {
            // Load the source DOCX document.
            Document doc = new Document("Input.docx");

            // Base name for the output files.
            string baseFileName = "Chapter";

            // Configure HTML save options.
            HtmlSaveOptions options = new HtmlSaveOptions();

            // Split the document at heading paragraphs (each chapter).
            options.DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph;

            // Optional: limit splitting to a specific heading level (e.g., Heading 1).
            options.DocumentSplitHeadingLevel = 1;

            // Assign a callback to control the naming of each generated HTML part.
            options.DocumentPartSavingCallback = new ChapterPartRenamer(baseFileName, options.DocumentSplitCriteria);

            // Save the document. Aspose.Words will create multiple HTML files, one per chapter.
            doc.Save(baseFileName + ".html", options);
        }

        // Callback that renames each document part (chapter) during the save operation.
        private class ChapterPartRenamer : IDocumentPartSavingCallback
        {
            private readonly string _baseName;
            private readonly DocumentSplitCriteria _criteria;
            private int _count;

            public ChapterPartRenamer(string baseName, DocumentSplitCriteria criteria)
            {
                _baseName = baseName;
                _criteria = criteria;
            }

            void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
            {
                // Build a filename like "Chapter_Chapter1.html", "Chapter_Chapter2.html", etc.
                string partFileName = $"{_baseName}_Chapter{++_count}{Path.GetExtension(args.DocumentPartFileName)}";

                // Set the filename for the part.
                args.DocumentPartFileName = partFileName;

                // Alternatively, provide a custom stream for the part.
                args.DocumentPartStream = new FileStream(partFileName, FileMode.Create);
            }
        }
    }
}
