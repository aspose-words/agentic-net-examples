using System;
using System.Collections.Generic;
using Aspose.Words;

namespace AsposeWordsSelectExample
{
    // DTO that holds paragraph text and its style name.
    public class ParagraphDto
    {
        // Initialise with empty strings to satisfy non‑nullable warnings.
        public string Text { get; set; } = string.Empty;
        public string StyleName { get; set; } = string.Empty;
    }

    class Program
    {
        static void Main()
        {
            // Load the DOTX template. The constructor is the approved lifecycle rule.
            Document doc = new Document("Template.dotx");

            // Retrieve all paragraph nodes from the document.
            NodeCollection paragraphNodes = doc.GetChildNodes(NodeType.Paragraph, true);

            // Project the selected nodes into a list of DTOs.
            List<ParagraphDto> paragraphs = new List<ParagraphDto>();
            foreach (Node node in paragraphNodes)
            {
                // Cast the node to Paragraph to access its properties.
                Paragraph para = node as Paragraph;
                if (para == null) continue;

                // Extract the paragraph text (trim the trailing paragraph break).
                string text = para.GetText().TrimEnd('\r', '\n');

                // Retrieve the style name applied to the paragraph.
                string styleName = para.ParagraphFormat.StyleName ?? string.Empty;

                paragraphs.Add(new ParagraphDto
                {
                    Text = text,
                    StyleName = styleName
                });
            }

            // Example usage: write the DTO data to the console.
            foreach (var dto in paragraphs)
            {
                Console.WriteLine($"Text: \"{dto.Text}\", Style: \"{dto.StyleName}\"");
            }

            // Save the (unchanged) document to demonstrate the approved save rule.
            doc.Save("Result.docx");
        }
    }
}
