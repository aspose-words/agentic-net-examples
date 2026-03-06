// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class ListFormatter
{
    static void Main()
    {
        // Load an existing DOCX document.
        // (Replace with the actual path to your document.)
        Document doc = new Document("Input.docx");

        // ------------------------------------------------------------
        // 1. Format list items in the main body.
        //    Replace every occurrence of the word "Item:" with a bullet character.
        // ------------------------------------------------------------
        FindReplaceOptions listOptions = new FindReplaceOptions();
        // Preserve the original formatting of the paragraph.
        listOptions.ApplyParagraphFormat = true;
        // Replace "Item:" with a bullet (Unicode U+2022) followed by a space.
        doc.Range.Replace("Item:", "\u2022 ", listOptions);

        // ------------------------------------------------------------
        // 2. Extract text between two bookmarks named "Start" and "End".
        //    This demonstrates working with node ranges.
        // ------------------------------------------------------------
        // Locate the bookmark start and end nodes.
        Node startNode = doc.Range.Bookmarks["Start"]?.BookmarkStart?.ParentNode;
        Node endNode = doc.Range.Bookmarks["End"]?.BookmarkEnd?.ParentNode;

        if (startNode != null && endNode != null)
        {
            // Create a temporary range that spans from the start node to the end node.
            // The Range property of a node gives access to the text it contains.
            // We'll concatenate the text of all nodes between the two bookmarks.
            string extractedText = "";
            Node current = startNode;
            while (current != null && current != endNode)
            {
                extractedText += current.GetText();
                current = current.NextSibling;
            }
            // Include the end node's text as well.
            extractedText += endNode.GetText();

            // For demonstration, write the extracted text to the console.
            Console.WriteLine("Extracted text between bookmarks:");
            Console.WriteLine(extractedText.Trim());
        }

        // ------------------------------------------------------------
        // 3. Process headers and footers in each section.
        //    Replace a placeholder "(CompanyName)" with an actual name.
        // ------------------------------------------------------------
        foreach (Section section in doc.Sections)
        {
            foreach (HeaderFooter headerFooter in section.HeadersFooters)
            {
                // Use the Range of the header/footer to perform a replace.
                headerFooter.Range.Replace("(CompanyName)", "Acme Corp.");
            }
        }

        // ------------------------------------------------------------
        // 4. Process footnotes (and endnotes) in the document.
        //    Replace a placeholder "[Note]" with "See footnote".
        // ------------------------------------------------------------
        NodeCollection footnotes = doc.GetChildNodes(NodeType.Footnote, true);
        foreach (Footnote footnote in footnotes)
        {
            footnote.Range.Replace("[Note]", "See footnote");
        }

        // ------------------------------------------------------------
        // 5. Save the modified document.
        // ------------------------------------------------------------
        doc.Save("Output.docx");
    }
}
