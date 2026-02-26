using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsExtractBetweenNodes
{
    class Program
    {
        static void Main()
        {
            // Load the source DOCX file.
            Document srcDoc = new Document("input.docx");

            // Example: extract all content that lies between the 2nd and 5th paragraphs (inclusive start, exclusive end).
            // Adjust the indices as needed for your specific scenario.
            Paragraph startParagraph = (Paragraph)srcDoc.GetChild(NodeType.Paragraph, 1, true); // 0‑based index
            Paragraph endParagraph   = (Paragraph)srcDoc.GetChild(NodeType.Paragraph, 4, true);

            // Collect the text of every node that appears after the start node and before the end node.
            StringBuilder extractedText = new StringBuilder();
            Node curNode = startParagraph.NextSibling;

            while (curNode != null && curNode != endParagraph)
            {
                // GetText() returns the text of the node and its children.
                extractedText.Append(curNode.GetText());
                curNode = curNode.NextSibling;
            }

            // Output the extracted text to the console.
            Console.WriteLine("Extracted content between the specified nodes:");
            Console.WriteLine(extractedText.ToString());

            // Optionally, create a new document that contains only the extracted content.
            Document resultDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(resultDoc);
            builder.Writeln(extractedText.ToString());

            // Save the new document.
            resultDoc.Save("extracted_content.docx");
        }
    }
}
