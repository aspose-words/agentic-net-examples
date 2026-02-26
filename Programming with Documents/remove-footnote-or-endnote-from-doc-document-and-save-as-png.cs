using System;
using Aspose.Words;
using Aspose.Words.Notes;
using Aspose.Words.Saving;

class RemoveNotesAndSaveAsPng
{
    static void Main()
    {
        // Load the existing Word document.
        Document doc = new Document("Input.docx");

        // Remove all footnotes.
        foreach (Footnote footnote in doc.GetChildNodes(NodeType.Footnote, true))
        {
            // Keep only footnotes (skip endnotes in this loop).
            if (footnote.FootnoteType == FootnoteType.Footnote)
                footnote.Remove();
        }

        // Remove all endnotes.
        // Endnotes are also returned by GetChildNodes(NodeType.Footnote, true) – they are just Footnote objects with type Endnote.
        foreach (Footnote note in doc.GetChildNodes(NodeType.Footnote, true))
        {
            if (note.FootnoteType == FootnoteType.Endnote)
                note.Remove();
        }

        // Configure image save options to render the document as PNG.
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
        {
            // Render only the first page (optional – remove to render all pages).
            PageSet = new PageSet(0)
        };

        // Save the modified document as a PNG image.
        doc.Save("Output.png", pngOptions);
    }
}
