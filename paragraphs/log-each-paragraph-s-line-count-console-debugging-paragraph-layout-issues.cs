using System;
using Aspose.Words;
using Aspose.Words.Layout;

class ParagraphLineCounter
{
    static void Main()
    {
        // Load an existing document.
        Document doc = new Document("Input.docx");

        // Ensure the layout model is up‑to‑date so we can query layout entities.
        doc.UpdatePageLayout();

        // Helper that maps nodes to their layout entities.
        LayoutCollector layoutCollector = new LayoutCollector(doc);
        // Enumerator used to walk through layout entities.
        LayoutEnumerator layoutEnumerator = new LayoutEnumerator(doc);

        // Iterate over every paragraph in the document.
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        for (int i = 0; i < paragraphs.Count; i++)
        {
            Paragraph paragraph = (Paragraph)paragraphs[i];

            // Obtain the layout entity that corresponds to this paragraph.
            object paragraphEntity = layoutCollector.GetEntity(paragraph);
            if (paragraphEntity == null)
            {
                Console.WriteLine($"Paragraph {i} has no layout entity.");
                continue;
            }

            // Position the enumerator on the paragraph entity.
            layoutEnumerator.Current = paragraphEntity;

            int lineCount = 0;
            bool scanningLineForRealText = false;

            // Walk through the children of the paragraph entity.
            if (layoutEnumerator.MoveFirstChild())
            {
                do
                {
                    // When we encounter a line entity we start looking for real text inside it.
                    if (layoutEnumerator.Type == LayoutEntityType.Line)
                        scanningLineForRealText = true;

                    // If the current entity is a span that contains text and we are inside a line,
                    // count this line as having real text.
                    if (scanningLineForRealText && layoutEnumerator.Type == LayoutEntityType.Span &&
                        layoutEnumerator.Kind.StartsWith("TEXT"))
                    {
                        lineCount++;
                        scanningLineForRealText = false; // reset for the next line
                    }
                }
                while (layoutEnumerator.MoveNext());

                // Return the enumerator to the paragraph entity for the next iteration.
                layoutEnumerator.Current = paragraphEntity;
            }

            Console.WriteLine($"Paragraph {i} contains {lineCount} line(s).");
        }

        // Save the (unchanged) document – this demonstrates the required save lifecycle.
        doc.Save("Output.docx");
    }
}
