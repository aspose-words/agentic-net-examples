using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Notes;

class RemoveFootnoteAndSaveAsJpeg
{
    static void Main()
    {
        // Load the source DOC/DOCX document.
        Document doc = new Document("InputDocument.docx");

        // Remove the footnote separator (the line that separates footnotes from the main text).
        // This effectively eliminates the visual footnote area.
        FootnoteSeparator footnoteSeparator = doc.FootnoteSeparators[FootnoteSeparatorType.FootnoteSeparator];
        footnoteSeparator.FirstParagraph.FirstChild?.Remove();

        // Remove the endnote separator in the same way (optional, if the document contains endnotes).
        FootnoteSeparator endnoteSeparator = doc.FootnoteSeparators[FootnoteSeparatorType.EndnoteSeparator];
        endnoteSeparator.FirstParagraph.FirstChild?.Remove();

        // Configure image save options for JPEG output.
        ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
        {
            // Render only the first page; JPEG format saves a single page.
            PageSet = new PageSet(0)
        };

        // Save the modified document as a JPEG image.
        doc.Save("OutputImage.jpg", jpegOptions);
    }
}
