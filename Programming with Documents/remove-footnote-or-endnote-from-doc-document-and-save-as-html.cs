using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Notes;

class RemoveNotesAndSaveHtml
{
    static void Main()
    {
        // Load the source DOC/DOCX document.
        Document doc = new Document("InputDocument.docx");

        // Remove footnote separator (if present) – this effectively removes footnote content from the output.
        FootnoteSeparator footnoteSeparator = doc.FootnoteSeparators[FootnoteSeparatorType.FootnoteSeparator];
        if (footnoteSeparator?.FirstParagraph?.FirstChild != null)
            footnoteSeparator.FirstParagraph.FirstChild.Remove();

        // Remove endnote separator (if present) – this effectively removes endnote content from the output.
        FootnoteSeparator endnoteSeparator = doc.FootnoteSeparators[FootnoteSeparatorType.EndnoteSeparator];
        if (endnoteSeparator?.FirstParagraph?.FirstChild != null)
            endnoteSeparator.FirstParagraph.FirstChild.Remove();

        // Configure HTML save options – omit headers/footers for a cleaner HTML output.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            ExportHeadersFootersMode = ExportHeadersFootersMode.None
        };

        // Save the modified document as HTML.
        doc.Save("OutputDocument.html", saveOptions);
    }
}
