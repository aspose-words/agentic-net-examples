using System;
using Aspose.Words;
using Aspose.Words.Saving;

class AppendFootnotesAndConvertToPdf
{
    static void Main()
    {
        // Load the destination document (contains existing footnotes).
        Document dstDoc = new Document("Destination.docx");

        // Load the source document (contains footnotes to be appended).
        Document srcDoc = new Document("SourceWithFootnotes.docx");

        // Count footnotes already present in the destination document.
        int existingFootnoteCount = dstDoc.GetChildNodes(NodeType.Footnote, true).Count;

        // Set the starting number for the source document's footnotes so that they continue after the existing ones.
        srcDoc.FootnoteOptions.StartNumber = existingFootnoteCount + 1;

        // Append the source document to the destination document.
        dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

        // Save the combined document as PDF.
        dstDoc.Save("CombinedResult.pdf", SaveFormat.Pdf);
    }
}
