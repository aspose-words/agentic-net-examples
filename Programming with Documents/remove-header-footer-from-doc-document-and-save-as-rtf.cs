using System;
using Aspose.Words;
using Aspose.Words.Saving;

class RemoveHeadersFootersAndSaveRtf
{
    static void Main()
    {
        // Path to the source DOC file.
        string inputPath = @"C:\Docs\SourceDocument.doc";

        // Path where the resulting RTF file will be saved.
        string outputPath = @"C:\Docs\ResultDocument.rtf";

        // Load the existing DOC document.
        Document doc = new Document(inputPath);

        // Remove headers and footers from every section in the document.
        foreach (Section section in doc.Sections)
        {
            // Clears the text of all header/footer types while keeping the objects.
            section.ClearHeadersFooters();
        }

        // Create RTF save options (default constructor).
        RtfSaveOptions rtfOptions = new RtfSaveOptions();

        // Save the modified document as RTF using the specified options.
        doc.Save(outputPath, rtfOptions);
    }
}
