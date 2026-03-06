using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOC/DOCX document.
        Document doc = new Document("Input.docx");

        // Remove all footnotes and endnotes from the document.
        // The XPath "//Footnote" selects every footnote/endnote node.
        NodeList footnoteNodes = doc.SelectNodes("//Footnote");
        foreach (Node footnote in footnoteNodes)
        {
            footnote.Remove();
        }

        // (Optional) Remove the footnote separator paragraph to avoid an empty line.
        // FootnoteSeparator separator = doc.FootnoteSeparators[FootnoteSeparatorType.FootnoteSeparator];
        // separator.FirstParagraph?.FirstChild?.Remove();

        // Configure image save options to render the document as a JPEG.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Jpeg);
        // Render only the first page (zero‑based index) – adjust as needed.
        saveOptions.PageSet = new PageSet(0);

        // Save the modified document as a JPEG image.
        doc.Save("Output.jpg", saveOptions);
    }
}
