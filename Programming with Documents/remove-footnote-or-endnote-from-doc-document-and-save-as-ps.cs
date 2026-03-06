using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Notes;
using Aspose.Words.Saving;

class RemoveNotesAndSaveAsPs
{
    static void Main()
    {
        // Path to the source DOC document.
        string dataDir = @"C:\Docs\";
        string inputPath = Path.Combine(dataDir, "Input.docx");

        // Load the document.
        Document doc = new Document(inputPath);

        // Remove the footnote separator (if it exists).
        FootnoteSeparator footnoteSeparator = doc.FootnoteSeparators[FootnoteSeparatorType.FootnoteSeparator];
        if (footnoteSeparator?.FirstParagraph?.FirstChild != null)
            footnoteSeparator.FirstParagraph.FirstChild.Remove();

        // Remove the endnote separator (if it exists).
        FootnoteSeparator endnoteSeparator = doc.FootnoteSeparators[FootnoteSeparatorType.EndnoteSeparator];
        if (endnoteSeparator?.FirstParagraph?.FirstChild != null)
            endnoteSeparator.FirstParagraph.FirstChild.Remove();

        // Configure PostScript save options.
        PsSaveOptions saveOptions = new PsSaveOptions
        {
            SaveFormat = SaveFormat.Ps
        };

        // Save the modified document as a PostScript file.
        string outputPath = Path.Combine(dataDir, "Output.ps");
        doc.Save(outputPath, saveOptions);
    }
}
