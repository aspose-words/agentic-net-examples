using System;
using System.Collections.Generic;
using Aspose.Words;

class StyleSeparatorFinder
{
    static void Main()
    {
        // Load the document (replace with your actual file path).
        Document doc = new Document("input.docx");

        // List to hold paragraphs that are style separators.
        List<Paragraph> styleSeparatorParagraphs = new List<Paragraph>();

        // Iterate through all paragraphs in the document.
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            // A style separator is represented by a paragraph break whose
            // BreakIsStyleSeparator property is true.
            if (para.BreakIsStyleSeparator)
                styleSeparatorParagraphs.Add(para);
        }

        // Example processing: output the number of found style separators.
        Console.WriteLine($"Found {styleSeparatorParagraphs.Count} style separator(s) in the document.");

        // (Optional) Further processing can be done here, e.g., modifying the paragraphs.

        // Save the document after any modifications.
        doc.Save("output.docx");
    }
}
