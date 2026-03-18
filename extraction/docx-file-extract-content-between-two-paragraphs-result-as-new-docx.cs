using System;
using Aspose.Words;

class ExtractBetweenParagraphs
{
    static void Main()
    {
        // Paths to the source and destination documents.
        const string inputPath = "input.docx";
        const string outputPath = "output.docx";

        // Text of the paragraphs that mark the start and end of the range to extract.
        const string startMarker = "Start of extraction";
        const string endMarker = "End of extraction";

        // Load the source document.
        Document srcDoc = new Document(inputPath);

        // Get all paragraphs in the document.
        NodeCollection allParagraphs = srcDoc.GetChildNodes(NodeType.Paragraph, true);

        // Locate the indices of the start and end markers.
        int startIdx = -1;
        int endIdx = -1;
        for (int i = 0; i < allParagraphs.Count; i++)
        {
            Paragraph para = (Paragraph)allParagraphs[i];
            string text = para.GetText().Trim(); // Trim to ignore trailing paragraph break.

            if (startIdx == -1 && text == startMarker)
                startIdx = i;

            if (text == endMarker)
            {
                endIdx = i;
                break; // Stop after finding the end marker.
            }
        }

        // Validate that both markers were found and that the range is logical.
        if (startIdx == -1 || endIdx == -1 || endIdx <= startIdx)
            throw new InvalidOperationException("Unable to locate valid start/end paragraphs.");

        // Create a new blank document.
        Document destDoc = new Document();

        // Prepare an importer to copy nodes while preserving formatting.
        NodeImporter importer = new NodeImporter(srcDoc, destDoc, ImportFormatMode.KeepSourceFormatting);

        // Import each paragraph that lies between the two markers (exclusive).
        for (int i = startIdx + 1; i < endIdx; i++)
        {
            Node importedNode = importer.ImportNode(allParagraphs[i], true);
            destDoc.FirstSection.Body.AppendChild(importedNode);
        }

        // Save the extracted content as a new DOCX file.
        destDoc.Save(outputPath);
    }
}
