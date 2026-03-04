using System;
using System.Collections.Generic;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the DOTM (macro‑enabled) document.
        Document doc = new Document("InputTemplate.dotm");

        // Store all heading texts.
        List<string> headings = new List<string>();

        // Retrieve every paragraph in the document (including those in headers/footers).
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        foreach (Paragraph para in paragraphs)
        {
            // ParagraphFormat.IsHeading is true for built‑in heading styles (Heading 1, Heading 2, …).
            if (para.ParagraphFormat.IsHeading)
            {
                // GetText() returns the paragraph text together with its child nodes.
                headings.Add(para.GetText().Trim());
            }
        }

        // Output the collected headings.
        foreach (string heading in headings)
        {
            Console.WriteLine(heading);
        }

        // Save the document if needed (no modifications were made).
        doc.Save("Output.docx");
    }
}
