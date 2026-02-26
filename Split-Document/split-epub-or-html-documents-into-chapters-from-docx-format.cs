using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class SplitDocumentExample
{
    static void Main()
    {
        // Load the source DOCX document.
        Document doc = new Document("Input.docx");

        // ---------- Save as EPUB with chapters ----------
        // Create HtmlSaveOptions for EPUB format.
        HtmlSaveOptions epubOptions = new HtmlSaveOptions(SaveFormat.Epub);
        // Split the document at heading paragraphs (e.g., Heading 1, Heading 2).
        epubOptions.DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph;
        // Define up to which heading level the split should occur.
        epubOptions.DocumentSplitHeadingLevel = 2; // split at Heading 1 and Heading 2.
        // Export document properties (optional).
        epubOptions.ExportDocumentProperties = true;
        // Assign a callback to control the filenames of each chapter part.
        epubOptions.DocumentPartSavingCallback = new ChapterPartRenamer("Chapter", DocumentSplitCriteria.HeadingParagraph);

        // Save the document as an EPUB file. Each chapter becomes a separate HTML part inside the EPUB.
        doc.Save("Output.epub", epubOptions);

        // ---------- Save as HTML with separate chapter files ----------
        // Create HtmlSaveOptions for HTML format.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
        htmlOptions.DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph;
        htmlOptions.DocumentSplitHeadingLevel = 2;
        htmlOptions.ExportDocumentProperties = true;
        htmlOptions.DocumentPartSavingCallback = new ChapterPartRenamer("Chapter", DocumentSplitCriteria.HeadingParagraph);
        // Optional: specify a folder where linked images will be stored.
        htmlOptions.ImagesFolder = "OutputImages";

        // Save the document as HTML. Multiple HTML files will be created, one per chapter.
        doc.Save("Output.html", htmlOptions);
    }

    // Callback implementation to rename each split document part.
    private class ChapterPartRenamer : IDocumentPartSavingCallback
    {
        private readonly string _baseName;
        private readonly DocumentSplitCriteria _criteria;
        private int _counter;

        public ChapterPartRenamer(string baseName, DocumentSplitCriteria criteria)
        {
            _baseName = baseName;
            _criteria = criteria;
            _counter = 0;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // Preserve the original file extension (e.g., .html or .xhtml).
            string extension = Path.GetExtension(args.DocumentPartFileName);
            // Build a new filename like "Chapter part 1.html", "Chapter part 2.html", etc.
            string newFileName = $"{_baseName} part {++_counter}{extension}";
            args.DocumentPartFileName = newFileName;
            // No custom stream is provided; Aspose.Words will handle the file creation.
        }
    }
}
