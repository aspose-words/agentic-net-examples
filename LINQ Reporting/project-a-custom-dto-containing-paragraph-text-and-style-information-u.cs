using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Markup;

namespace AsposeWordsExample
{
    // DTO that holds paragraph text and its style name.
    public class ParagraphDto
    {
        public string Text { get; set; }
        public string StyleName { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Load an existing DOC/DOCX document.
            Document doc = new Document("InputDocument.docx");   // create/load rule

            // Use XPath to select all paragraph nodes.
            // The XPath expression selects every paragraph element in the document.
            // Aspose.Words automatically handles the WordML namespace for simple queries.
            NodeList paragraphNodes = doc.SelectNodes("//Paragraph");   // select rule

            // Project the selected nodes into a list of DTOs.
            List<ParagraphDto> paragraphs = new List<ParagraphDto>();
            foreach (Paragraph para in paragraphNodes)
            {
                // Get the visible text of the paragraph (including the paragraph break).
                string text = para.GetText().TrimEnd('\r', '\a');

                // Retrieve the style name applied to the paragraph.
                string styleName = para.ParagraphFormat.StyleName;

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

            // Optionally, create a new document that lists the extracted information.
            Document outputDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(outputDoc);   // create rule

            foreach (var dto in paragraphs)
            {
                builder.Writeln($"Paragraph Text: {dto.Text}");
                builder.Writeln($"Style Name: {dto.StyleName}");
                builder.Writeln(); // add an empty line between entries
            }

            // Save the result document.
            outputDoc.Save("ExtractedParagraphs.docx");   // save rule
        }
    }
}
