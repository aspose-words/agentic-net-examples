// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using Aspose.Words;
using Aspose.Words.Saving;

class RemoveNotesAndSaveAsMhtml
{
    static void Main()
    {
        // Path to the source DOC/DOCX file.
        string inputPath = @"C:\Docs\SourceDocument.docx";

        // Path where the resulting MHTML file will be saved.
        string outputPath = @"C:\Docs\ResultDocument.mhtml";

        // Load the document from the file system.
        Document doc = new Document(inputPath);

        // Remove all footnotes.
        NodeCollection footnotes = doc.GetChildNodes(NodeType.Footnote, true);
        for (int i = footnotes.Count - 1; i >= 0; i--)
        {
            footnotes[i].Remove();
        }

        // Remove all endnotes.
        NodeCollection endnotes = doc.GetChildNodes(NodeType.Endnote, true);
        for (int i = endnotes.Count - 1; i >= 0; i--)
        {
            endnotes[i].Remove();
        }

        // Save the modified document as MHTML.
        doc.Save(outputPath, SaveFormat.Mhtml);
    }
}
