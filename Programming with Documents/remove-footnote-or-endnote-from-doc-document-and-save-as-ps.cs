using Aspose.Words;
using Aspose.Words.Notes;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOC document.
        Document doc = new Document("Input.doc");

        // Remove the footnote separator line, if it exists.
        FootnoteSeparator footnoteSeparator = doc.FootnoteSeparators[FootnoteSeparatorType.FootnoteSeparator];
        if (footnoteSeparator.FirstParagraph?.FirstChild != null)
            footnoteSeparator.FirstParagraph.FirstChild.Remove();

        // Remove the endnote separator line, if it exists.
        FootnoteSeparator endnoteSeparator = doc.FootnoteSeparators[FootnoteSeparatorType.EndnoteSeparator];
        if (endnoteSeparator.FirstParagraph?.FirstChild != null)
            endnoteSeparator.FirstParagraph.FirstChild.Remove();

        // Configure PostScript save options.
        PsSaveOptions saveOptions = new PsSaveOptions
        {
            SaveFormat = SaveFormat.Ps
        };

        // Save the modified document as a PostScript file.
        doc.Save("Output.ps", saveOptions);
    }
}
