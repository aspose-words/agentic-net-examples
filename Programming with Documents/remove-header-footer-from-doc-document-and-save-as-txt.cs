using System;
using Aspose.Words;
using Aspose.Words.Saving;

class RemoveHeadersFootersAndSaveAsTxt
{
    static void Main()
    {
        // Path to the source DOC document.
        string inputPath = @"C:\Docs\SourceDocument.docx";

        // Path for the resulting TXT file.
        string outputPath = @"C:\Docs\ResultDocument.txt";

        // Load the existing Word document.
        Document doc = new Document(inputPath);

        // Remove all header and footer contents from each section.
        foreach (Section section in doc.Sections)
        {
            // Clears the text inside headers/footers but keeps the objects.
            section.ClearHeadersFooters();
        }

        // Configure TXT save options to ensure no headers/footers are exported.
        TxtSaveOptions txtOptions = new TxtSaveOptions
        {
            ExportHeadersFootersMode = TxtExportHeadersFootersMode.None
        };

        // Save the document as plain text using the configured options.
        doc.Save(outputPath, txtOptions);
    }
}
