using System;
using Aspose.Words;

namespace InsertParagraphExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load the source document that contains the paragraph to be copied
            Document srcDoc = new Document("Source.doc");

            // Create a new blank destination document (it already contains one empty section)
            Document dstDoc = new Document();

            // Get the first (or any) section of the destination document where we will insert the paragraph
            Section destSection = dstDoc.Sections[0];

            // Retrieve the paragraph you want to copy from the source document
            // Here we take the first paragraph of the first section, adjust as needed
            Paragraph srcParagraph = srcDoc.FirstSection.Body.FirstParagraph;

            // Import the paragraph node into the destination document's node collection
            Node importedParagraph = dstDoc.ImportNode(srcParagraph, true);

            // Append the imported paragraph to the body of the destination section
            destSection.Body.AppendChild(importedParagraph);

            // Save the resulting document as a DOC file
            dstDoc.Save("Result.doc");
        }
    }
}
