using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables; // For NodeType enum

class Program
{
    static void Main()
    {
        // Load the existing Word document.
        // This uses the Document(string) constructor, which is the provided load rule.
        Document doc = new Document("Input.docx");

        // Retrieve all Paragraph nodes in the document (including those in headers/footers if needed).
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true)
                           .Cast<Paragraph>()
                           .Select(p => p.GetText().Trim())
                           .ToList();

        // Output each paragraph text to the console.
        foreach (var text in paragraphs)
        {
            Console.WriteLine(text);
        }

        // Optionally, save the document after processing (uses the Document.Save(string) method).
        // doc.Save("Output.docx");
    }
}
