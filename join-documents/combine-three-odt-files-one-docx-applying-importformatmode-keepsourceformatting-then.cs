using System;
using Aspose.Words;

class OdtMerger
{
    static void Main()
    {
        // Paths to the source ODT files.
        string[] sourceFiles = {
            "Document1.odt",
            "Document2.odt",
            "Document3.odt"
        };

        // Create a blank destination document.
        Document destination = new Document();

        // Load each ODT file and append it to the destination document,
        // preserving the original formatting of each source.
        foreach (string filePath in sourceFiles)
        {
            Document source = new Document(filePath);                     // Load ODT
            destination.AppendDocument(source, ImportFormatMode.KeepSourceFormatting); // Append with formatting
        }

        // Save the combined document as DOCX. The format is inferred from the extension.
        destination.Save("Combined.docx");
    }
}
