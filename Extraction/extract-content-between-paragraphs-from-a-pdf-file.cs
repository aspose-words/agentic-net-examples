using System;
using Aspose.Words;

namespace ExtractContentBetweenParagraphs
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source PDF file
            string inputPdfPath = @"C:\Docs\source.pdf";

            // Load the PDF file as an Aspose.Words document. Aspose.Words can open PDF directly.
            Document sourceDoc = new Document(inputPdfPath);

            // Get all paragraphs in the main body of the first section
            ParagraphCollection allParagraphs = sourceDoc.FirstSection.Body.Paragraphs;

            // Define the unique text that marks the start and end paragraphs
            string startMarker = "=== Start ===";
            string endMarker   = "=== End ===";

            // Locate the indices of the start and end markers
            int startIndex = -1;
            int endIndex   = -1;

            for (int i = 0; i < allParagraphs.Count; i++)
            {
                string paragraphText = allParagraphs[i].GetText();

                if (startIndex == -1 && paragraphText.Contains(startMarker))
                    startIndex = i;

                if (paragraphText.Contains(endMarker))
                    endIndex = i;
            }

            // Ensure valid indices were found and that there is content between them
            if (startIndex != -1 && endIndex != -1 && endIndex > startIndex + 1)
            {
                // Create a new blank document to hold the extracted content
                Document extractedDoc = new Document();

                // Build the minimal required structure (Section -> Body)
                Section newSection = new Section(extractedDoc);
                extractedDoc.AppendChild(newSection);
                Body newBody = new Body(extractedDoc);
                newSection.AppendChild(newBody);

                // Import each paragraph that lies between the markers
                for (int i = startIndex + 1; i < endIndex; i++)
                {
                    // ImportNode clones the node into the target document while preserving formatting
                    Node importedParagraph = extractedDoc.ImportNode(allParagraphs[i], true);
                    newBody.AppendChild(importedParagraph);
                }

                // Save the extracted content to a new DOCX file (or any format supported by the extension)
                string outputDocxPath = @"C:\Docs\extracted_content.docx";
                extractedDoc.Save(outputDocxPath);
                Console.WriteLine($"Extracted content saved to: {outputDocxPath}");
            }
            else
            {
                // No valid range found – handle as needed (e.g., log, throw, etc.)
                Console.WriteLine("Unable to locate a valid start/end paragraph range.");
            }
        }
    }
}
