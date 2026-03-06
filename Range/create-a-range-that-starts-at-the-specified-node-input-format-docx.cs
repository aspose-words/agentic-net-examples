using System;
using Aspose.Words;
using AsposeRange = Aspose.Words.Range; // Alias to avoid conflict with System.Range

namespace AsposeWordsRangeExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Input and output file paths (replace with actual paths as needed)
            string inputPath = @"InputDocument.docx";
            string outputPath = @"OutputDocument.docx";

            // Load the DOCX document using Aspose.Words Document constructor (create/load lifecycle)
            Document doc = new Document(inputPath);

            // Obtain a node from which we want to start the range.
            // For this example we use the first paragraph in the first section.
            Node startNode = doc.FirstSection.Body.FirstParagraph;

            // Get the Range that starts at the specified node.
            // The Range object represents a contiguous area of the document that includes the start node.
            AsposeRange range = startNode.Range; // Explicitly use the Aspose.Words.Range type

            // Example operation on the range: output its text to the console.
            Console.WriteLine("Range text starts with:");
            Console.WriteLine(range.Text.Trim());

            // Optionally, modify the range (e.g., replace a word) – this demonstrates that the range is functional.
            // Replace the first occurrence of "Hello" with "Hi" within the range.
            if (range.Text.Contains("Hello"))
            {
                range.Replace("Hello", "Hi");
            }

            // Save the modified document using the Document.Save method (save lifecycle)
            doc.Save(outputPath);
        }
    }
}
