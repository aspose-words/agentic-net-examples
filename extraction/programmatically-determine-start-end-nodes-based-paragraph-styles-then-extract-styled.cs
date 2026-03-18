using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Tables;

class StyledContentExtractor
{
    static void Main()
    {
        // Path to the source document.
        string inputPath = "input.docx";

        // Path where the extracted styled segment will be saved as plain text.
        string extractedTextPath = "styledSegment.txt";

        // Load the document (lifecycle rule: load).
        Document doc = new Document(inputPath);

        // The style name we want to search for.
        string targetStyle = "Heading 1";

        // Get all paragraphs in the document (including those in headers/footers).
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

        int firstIndex = -1;
        int lastIndex = -1;

        // Determine the first and last paragraph indices that use the target style.
        for (int i = 0; i < paragraphs.Count; i++)
        {
            Paragraph para = (Paragraph)paragraphs[i];
            if (para.ParagraphFormat.StyleName == targetStyle)
            {
                if (firstIndex == -1)
                    firstIndex = i;          // First occurrence.
                lastIndex = i;               // Keep updating to the latest occurrence.
            }
        }

        if (firstIndex != -1)
        {
            // Build the text of the segment that starts with the first styled paragraph
            // and ends with the last styled paragraph (inclusive).
            StringBuilder segmentBuilder = new StringBuilder();

            for (int i = firstIndex; i <= lastIndex; i++)
            {
                Paragraph para = (Paragraph)paragraphs[i];
                // GetText returns the paragraph text plus the paragraph break character.
                segmentBuilder.Append(para.GetText());
            }

            // Write the extracted segment to a plain‑text file.
            File.WriteAllText(extractedTextPath, segmentBuilder.ToString());
        }
        else
        {
            Console.WriteLine($"No paragraphs with style '{targetStyle}' were found.");
        }

        // Save the (unchanged) document if needed (lifecycle rule: save).
        doc.Save("output.docx");
    }
}
