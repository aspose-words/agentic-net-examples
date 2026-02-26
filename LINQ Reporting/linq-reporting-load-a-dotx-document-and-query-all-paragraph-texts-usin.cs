using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables; // Not required but included for completeness

namespace AsposeWordsLinqExample
{
    class Program
    {
        static void Main()
        {
            // Load the DOTX template document from the file system.
            // This uses the built‑in Document constructor (load rule) – no custom loading code.
            Document doc = new Document("Template.dotx");

            // Retrieve all Paragraph nodes in the document (including those inside tables, headers, etc.).
            // GetChildNodes returns a live collection; we cast it to Paragraph for LINQ processing.
            List<Paragraph> paragraphs = doc.GetChildNodes(NodeType.Paragraph, true)
                                            .Cast<Paragraph>()
                                            .ToList();

            // Use LINQ to project each Paragraph to its plain text representation.
            // GetText() returns the paragraph text including the end‑of‑paragraph marker,
            // so we Trim() to remove trailing whitespace and line breaks.
            IEnumerable<string> paragraphTexts = paragraphs
                                                    .Select(p => p.GetText().Trim())
                                                    .Where(text => !string.IsNullOrEmpty(text));

            // Output the collected paragraph texts to the console.
            Console.WriteLine("Paragraph texts found in the document:");
            foreach (string text in paragraphTexts)
            {
                Console.WriteLine("- " + text);
            }

            // (Optional) Save the document after any modifications if needed.
            // This uses the built‑in Save method (save rule) – no custom saving logic.
            // doc.Save("Output.docx");
        }
    }
}
