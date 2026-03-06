using System;
using System.Linq;
using System.Collections.Generic;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the DOCX file to be loaded.
        string filePath = "input.docx";

        // Load the document using the Document(string) constructor (provided load rule).
        Document doc = new Document(filePath);

        // Retrieve all Paragraph nodes in the document (deep search).
        // GetChildNodes returns a NodeCollection; cast each node to Paragraph.
        List<string> paragraphTexts = doc
            .GetChildNodes(NodeType.Paragraph, true)          // true => search all descendant nodes
            .Cast<Paragraph>()                               // LINQ cast to Paragraph type
            .Select(p => p.GetText().Trim())                 // Extract the text of each paragraph and trim trailing breaks
            .ToList();

        // Output each paragraph text to the console.
        foreach (string text in paragraphTexts)
        {
            Console.WriteLine(text);
        }
    }
}
