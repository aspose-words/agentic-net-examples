using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables; // Needed for Paragraph type if not directly in Aspose.Words

class Program
{
    static void Main()
    {
        // Load the DOCM document using the provided Document constructor (lifecycle rule).
        Document doc = new Document("InputDocument.docm");

        // Query all paragraph nodes in the document using LINQ.
        var paragraphTexts = doc
            .GetChildNodes(NodeType.Paragraph, true)               // Get all paragraphs (deep search).
            .Cast<Paragraph>()                                    // Cast to Paragraph type.
            .Select(p => p.GetText().Trim())                      // Extract and trim the text of each paragraph.
            .ToList();

        // Output the collected paragraph texts.
        foreach (var text in paragraphTexts)
        {
            Console.WriteLine(text);
        }
    }
}
