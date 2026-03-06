using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the DOTM template. The Document constructor handles loading from a file.
        Document doc = new Document(@"C:\Path\To\Template.dotm");

        // Retrieve all Paragraph nodes in the document (including those in headers/footers if needed).
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true)
                           .Cast<Paragraph>();

        // Use LINQ to project each Paragraph to its plain text (trimmed of trailing breaks).
        var paragraphTexts = paragraphs
                             .Select(p => p.GetText().Trim())
                             .ToList();

        // Output the collected paragraph texts.
        foreach (var text in paragraphTexts)
        {
            Console.WriteLine(text);
        }
    }
}
