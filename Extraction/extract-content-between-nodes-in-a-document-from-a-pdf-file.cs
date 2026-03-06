using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Tables;

namespace ExtractBetweenNodes
{
    class Program
    {
        static void Main()
        {
            // Load the source PDF file as a Word document.
            // The Document constructor handles loading and detects the PDF format.
            Document sourceDoc = new Document("input.pdf");

            // Define the marker texts that indicate the start and end of the region to extract.
            const string startMarker = "START_MARKER";
            const string endMarker   = "END_MARKER";

            // Locate the paragraphs that contain the start and end markers.
            Paragraph startParagraph = null;
            Paragraph endParagraph   = null;

            // Iterate through all paragraphs in the main story.
            foreach (Paragraph para in sourceDoc.FirstSection.Body.Paragraphs)
            {
                string paraText = para.GetText();

                if (startParagraph == null && paraText.Contains(startMarker))
                    startParagraph = para;

                if (endParagraph == null && paraText.Contains(endMarker))
                    endParagraph = para;

                // Break early if both markers have been found.
                if (startParagraph != null && endParagraph != null)
                    break;
            }

            // Validate that both markers were found.
            if (startParagraph == null || endParagraph == null)
                throw new InvalidOperationException("Start or end marker not found in the document.");

            // Build the extracted text by traversing nodes between the two markers.
            StringBuilder extractedBuilder = new StringBuilder();

            // Start with the node immediately after the start marker.
            Node currentNode = startParagraph.NextSibling;

            // Continue until we reach the end marker node.
            while (currentNode != null && currentNode != endParagraph)
            {
                // Append the text of each node (paragraphs, tables, etc.).
                extractedBuilder.Append(currentNode.GetText());
                currentNode = currentNode.NextSibling;
            }

            // Create a new blank document to hold the extracted content.
            Document extractedDoc = new Document();

            // Use DocumentBuilder to insert the extracted text.
            DocumentBuilder builder = new DocumentBuilder(extractedDoc);
            builder.Writeln(extractedBuilder.ToString());

            // Save the extracted content to a new file.
            extractedDoc.Save("extracted.docx");
        }
    }
}
