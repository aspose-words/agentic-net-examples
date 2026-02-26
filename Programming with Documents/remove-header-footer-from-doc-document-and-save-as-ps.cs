using System;
using Aspose.Words;
using Aspose.Words.Saving;

class RemoveHeadersFootersAndSaveAsPs
{
    static void Main()
    {
        // Path to the source DOC document.
        string inputPath = @"C:\Docs\SourceDocument.doc";

        // Path where the resulting PostScript file will be saved.
        string outputPath = @"C:\Docs\ResultDocument.ps";

        // Load the existing DOC document.
        Document doc = new Document(inputPath);

        // Remove headers and footers from every section.
        foreach (Section section in doc.Sections)
        {
            // Clears the content of headers and footers while keeping the objects.
            section.ClearHeadersFooters();
        }

        // Configure save options for the PostScript format.
        PsSaveOptions saveOptions = new PsSaveOptions
        {
            // Explicitly set the format to PostScript (optional, but clear).
            SaveFormat = SaveFormat.Ps
        };

        // Save the modified document as a PostScript file.
        doc.Save(outputPath, saveOptions);
    }
}
