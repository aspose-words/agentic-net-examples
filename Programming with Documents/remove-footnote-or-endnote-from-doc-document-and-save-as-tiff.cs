using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Notes;

class RemoveNotesAndConvertToTiff
{
    static void Main()
    {
        // Path to the source DOC/DOCX file that contains footnotes/endnotes.
        string inputPath = @"C:\Docs\SourceDocument.docx";

        // Path where the resulting TIFF image will be saved.
        string outputPath = @"C:\Docs\ResultImage.tiff";

        // Load the document.
        Document doc = new Document(inputPath);

        // Remove all footnotes.
        NodeCollection footnotes = doc.GetChildNodes(NodeType.Footnote, true);
        foreach (Footnote footnote in footnotes)
        {
            // Only remove footnotes (skip endnotes if you want to keep them).
            if (footnote.FootnoteType == FootnoteType.Footnote)
                footnote.Remove();
        }

        // Remove all endnotes.
        NodeCollection endnotes = doc.GetChildNodes(NodeType.Footnote, true);
        foreach (Footnote endnote in endnotes)
        {
            if (endnote.FootnoteType == FootnoteType.Endnote)
                endnote.Remove();
        }

        // Configure image save options for TIFF format.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Optional: set compression, resolution, etc.
            TiffCompression = TiffCompression.Lzw,
            Resolution = 300
        };

        // Save the modified document as a TIFF image.
        doc.Save(outputPath, saveOptions);
    }
}
