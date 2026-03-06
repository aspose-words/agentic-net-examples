using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables; // For NodeType enum

class Program
{
    static void Main()
    {
        // Load the DOCM document from the file system.
        // The Document constructor is the provided load rule.
        Document doc = new Document(@"C:\Docs\Sample.docm");

        // Retrieve all Paragraph nodes in the document (including those in headers/footers).
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true)
                           .Cast<Paragraph>()
                           .Select(p => p.GetText().Trim())
                           .ToList();

        // Output each paragraph's text to the console.
        foreach (var text in paragraphs)
        {
            Console.WriteLine(text);
        }
    }
}
