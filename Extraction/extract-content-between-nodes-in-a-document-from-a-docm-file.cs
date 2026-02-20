using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load a DOCM file (the document may contain macros, Aspose.Words handles it automatically).
        // The Document constructor can open DOCM files without explicit LoadOptions.
        var doc = new Document(@"C:\Input\Sample.docm");

        // ------------------------------------------------------------
        // 1. Extraction using a pair of bookmarks named "Start" and "End".
        // ------------------------------------------------------------
        if (doc.Range.Bookmarks["Start"] != null && doc.Range.Bookmarks["End"] != null)
        {
            // If the "Start" bookmark spans the whole region we need, its Text property returns the content.
            // For separate start/end bookmarks you would need to locate the nodes manually (see the else block).
            string extractedText = doc.Range.Bookmarks["Start"].Text;

            Console.WriteLine("Extracted content between bookmarks:");
            Console.WriteLine(extractedText);
        }
        else
        {
            // ------------------------------------------------------------
            // 2. Extraction between arbitrary nodes (e.g., first and third paragraphs).
            // ------------------------------------------------------------
            Paragraph startParagraph = (Paragraph)doc.GetChild(NodeType.Paragraph, 0, true);
            Paragraph endParagraph   = (Paragraph)doc.GetChild(NodeType.Paragraph, 2, true);

            // Create a new empty document that will hold the extracted fragment.
            Document fragment = new Document();
            // Remove the default empty section that a new Document creates – we will reuse its first section.
            fragment.RemoveAllChildren();
            // Add a fresh empty section so we have a body to append nodes to.
            Section fragmentSection = new Section(doc);
            fragment.AppendChild(fragmentSection);
            Body fragmentBody = fragmentSection.Body;

            // Import the start node (including its children) into the fragment.
            Node importedStart = fragment.ImportNode(startParagraph, true);
            fragmentBody.AppendChild(importedStart);

            // Walk through sibling nodes until we reach the end node, importing each into the fragment.
            Node current = startParagraph.NextSibling;
            while (current != null && current != endParagraph)
            {
                Node imported = fragment.ImportNode(current, true);
                fragmentBody.AppendChild(imported);
                current = current.NextSibling;
            }

            // Finally, import the end node itself if you want it included.
            Node importedEnd = fragment.ImportNode(endParagraph, true);
            fragmentBody.AppendChild(importedEnd);

            // Retrieve the plain text of the extracted fragment.
            string extractedText = fragment.GetText();

            Console.WriteLine("Extracted content between specified paragraphs:");
            Console.WriteLine(extractedText);
        }
    }
}
