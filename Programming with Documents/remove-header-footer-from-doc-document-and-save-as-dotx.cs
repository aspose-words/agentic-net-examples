using System;
using Aspose.Words;
using Aspose.Words.Saving;

class RemoveHeadersFootersAndSaveAsTemplate
{
    static void Main()
    {
        // Path to the source DOC file.
        string sourcePath = @"C:\Docs\SourceDocument.doc";

        // Path where the resulting DOTX template will be saved.
        string outputPath = @"C:\Docs\ResultTemplate.dotx";

        // Load the existing DOC document.
        Document doc = new Document(sourcePath);

        // Iterate through all sections and clear their headers and footers.
        foreach (Section section in doc.Sections)
        {
            // This clears the text of all header/footer types in the section.
            section.ClearHeadersFooters();
        }

        // Save the modified document as a DOTX template.
        doc.Save(outputPath, SaveFormat.Dotx);
    }
}
