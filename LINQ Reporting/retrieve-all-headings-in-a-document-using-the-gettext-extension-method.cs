using System;
using System.Collections.Generic;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the DOTX template document.
        Document doc = new Document("Template.dotx");

        // Collect all heading texts.
        List<string> headings = new List<string>();

        // Iterate through all paragraphs in the document.
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            // Check if the paragraph style is a built‑in heading.
            if (para.ParagraphFormat.IsHeading)
            {
                // Get the paragraph text (including any child runs) and trim line breaks.
                string headingText = para.GetText().Trim();
                headings.Add(headingText);
            }
        }

        // Output the headings to the console.
        foreach (string heading in headings)
        {
            Console.WriteLine(heading);
        }

        // Save the document if needed (optional).
        doc.Save("Result.docx");
    }
}
