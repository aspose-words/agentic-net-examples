using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsExamples
{
    class JoinAdjacentTables
    {
        static void Main()
        {
            // Load the source document that contains adjacent tables.
            Document doc = new Document("Input.docx");

            // Get a live collection of all child nodes (paragraphs, tables, etc.) in the body of the first section.
            // Use GetChildNodes(NodeType.Any, true) because Body does not expose a ChildNodes property.
            NodeCollection bodyChildren = doc.FirstSection.Body.GetChildNodes(NodeType.Any, true);

            // Iterate through the collection using an index because we may modify the collection while iterating.
            for (int i = 0; i < bodyChildren.Count - 1; i++)
            {
                // Check if the current node is a table.
                if (bodyChildren[i].NodeType == NodeType.Table)
                {
                    // Check if the next node is also a table (adjacent tables).
                    if (bodyChildren[i + 1].NodeType == NodeType.Table)
                    {
                        // Create an empty paragraph.
                        Paragraph emptyParagraph = new Paragraph(doc);

                        // Insert the empty paragraph immediately after the current table.
                        // This will separate the two tables with a paragraph break.
                        doc.FirstSection.Body.InsertAfter(emptyParagraph, bodyChildren[i]);

                        // After insertion, the collection has grown by one node.
                        // Increment i to skip over the newly inserted paragraph.
                        i++;
                    }
                }
            }

            // Save the modified document.
            doc.Save("Output.docx");
        }
    }
}
