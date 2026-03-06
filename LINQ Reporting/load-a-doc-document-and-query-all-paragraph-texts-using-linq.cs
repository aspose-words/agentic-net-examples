using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Path to the Word document to load.
        string docPath = @"C:\Docs\Sample.docx";

        // Load the document using the Aspose.Words Document constructor.
        Document doc = new Document(docPath);

        // Retrieve all Paragraph nodes in the document (including those in headers/footers if needed).
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true)
                           .Cast<Paragraph>();

        // Use LINQ to project each paragraph to its trimmed text.
        var paragraphTexts = paragraphs
                             .Select(p => p.GetText().Trim())
                             .ToList();

        // Output the paragraph texts.
        foreach (var text in paragraphTexts)
        {
            Console.WriteLine(text);
        }
    }
}
