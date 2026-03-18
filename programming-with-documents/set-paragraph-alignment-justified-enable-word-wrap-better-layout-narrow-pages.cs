using System;
using Aspose.Words;
using Aspose.Words.Tables;

class ParagraphFormattingExample
{
    static void Main()
    {
        // Load an existing document.
        // Replace "input.docx" with the path to your source file.
        Document doc = new Document("input.docx");

        // Iterate over all paragraphs in the document.
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            // Set the paragraph alignment to justified.
            para.ParagraphFormat.Alignment = ParagraphAlignment.Justify;

            // Enable word wrap so that words are wrapped as whole units
            // (prevents mid‑word breaks on narrow pages).
            para.ParagraphFormat.WordWrap = true;
        }

        // Save the modified document.
        // Replace "output.docx" with the desired output path.
        doc.Save("output.docx");
    }
}
