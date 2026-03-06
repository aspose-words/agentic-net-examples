using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the DOCX document.
        Document doc = new Document("Input.docx");

        // Iterate through all paragraphs in the document.
        int paragraphIndex = 0;
        foreach (Paragraph paragraph in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            // Identify whether this paragraph break is a style separator.
            bool isStyleSeparator = paragraph.BreakIsStyleSeparator;

            // Output the result.
            Console.WriteLine($"Paragraph {paragraphIndex}: BreakIsStyleSeparator = {isStyleSeparator}");
            paragraphIndex++;
        }

        // Save the document (no modifications made, but required by lifecycle rules).
        doc.Save("Output.docx");
    }
}
