using System;
using Aspose.Words;
// Alias to avoid conflict with System.Range introduced in C# 8.0
using AsposeRange = Aspose.Words.Range;

class Program
{
    static void Main()
    {
        // Load the DOCX document (replace with your actual file path)
        Document doc = new Document("input.docx");

        // Specify the node from which the range should start.
        // Here we use the first paragraph of the first section as an example.
        Node startNode = doc.FirstSection.Body.FirstParagraph;

        // Create a Range that starts at the specified node.
        // Use the alias to disambiguate the type name.
        AsposeRange range = startNode.Range;

        // Example usage: output the text contained in the range.
        Console.WriteLine("Range text:");
        Console.WriteLine(range.Text);

        // Optionally, save the document after any modifications.
        doc.Save("output.docx");
    }
}
