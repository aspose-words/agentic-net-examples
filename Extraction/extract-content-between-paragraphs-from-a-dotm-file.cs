using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the DOTM file (macro‑enabled template)
        Document doc = new Document("Template.dotm");

        // Retrieve all paragraph nodes from the main story (including those in headers/footers if needed)
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

        // Build a string that contains the text of each paragraph,
        // separating consecutive paragraphs with a line break.
        StringBuilder extracted = new StringBuilder();

        for (int i = 0; i < paragraphs.Count; i++)
        {
            Paragraph para = (Paragraph)paragraphs[i];

            // GetText() returns the paragraph text including the end‑of‑paragraph character.
            // Trim the trailing paragraph break characters for clean output.
            string text = para.GetText().TrimEnd('\r', '\a');

            extracted.Append(text);

            // Add a line break between paragraphs, except after the last one.
            if (i < paragraphs.Count - 1)
                extracted.AppendLine();
        }

        // Save the extracted content to a plain‑text file.
        File.WriteAllText("ExtractedContent.txt", extracted.ToString());
    }
}
