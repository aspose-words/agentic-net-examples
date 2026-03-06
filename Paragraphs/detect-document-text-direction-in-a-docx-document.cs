using System;
using Aspose.Words;
using Aspose.Words.Loading;

class DetectDocumentDirection
{
    static void Main()
    {
        // Path to the DOCX file whose text direction we want to detect.
        string docPath = "input.docx";

        // Load the document. No special LoadOptions are required for DOCX.
        Document doc = new Document(docPath);

        // Iterate through all paragraphs in the document.
        foreach (Paragraph paragraph in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            // The Bidi property of ParagraphFormat is true for right‑to‑left paragraphs.
            bool isRightToLeft = paragraph.ParagraphFormat.Bidi;

            // Output the direction of the current paragraph.
            Console.WriteLine($"Paragraph #{paragraph.GetHashCode()}: {(isRightToLeft ? "Right‑to‑Left" : "Left‑to‑Right")}");
        }
    }
}
