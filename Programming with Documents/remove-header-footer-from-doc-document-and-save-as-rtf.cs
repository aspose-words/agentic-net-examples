using System;
using Aspose.Words;
using Aspose.Words.Saving;

class RemoveHeadersFootersAndSaveAsRtf
{
    static void Main()
    {
        // Path to the source DOC document.
        string inputPath = @"C:\Docs\Input.doc";

        // Path where the resulting RTF document will be saved.
        string outputPath = @"C:\Docs\Output.rtf";

        // Load the existing DOC document.
        Document doc = new Document(inputPath);

        // Remove all headers and footers from every section.
        foreach (Section section in doc.Sections)
        {
            section.ClearHeadersFooters(); // Clears the content but keeps the objects linked.
        }

        // Create RTF save options (default constructor).
        RtfSaveOptions rtfOptions = new RtfSaveOptions();

        // Save the modified document as RTF using the save options.
        doc.Save(outputPath, rtfOptions);
    }
}
