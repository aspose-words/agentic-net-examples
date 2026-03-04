using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables; // For NodeType enum if needed

class Program
{
    static void Main()
    {
        // Path to the DOT (template) document.
        string templatePath = @"C:\Docs\Template.dot";

        // Load the DOT document using the Document(string) constructor (load rule).
        Document doc = new Document(templatePath);

        // Query all Paragraph nodes in the document using LINQ.
        var paragraphTexts = doc
            .GetChildNodes(NodeType.Paragraph, true)               // Get all paragraph nodes (deep search).
            .Cast<Paragraph>()                                    // Cast to Paragraph type.
            .Select(p => p.GetText().Trim())                      // Extract and trim the text of each paragraph.
            .ToList();

        // Output the paragraph texts to the console.
        Console.WriteLine("Paragraphs found in the template:");
        foreach (var text in paragraphTexts)
        {
            Console.WriteLine($"- {text}");
        }

        // (Optional) Save the document to a new file to demonstrate the save rule.
        string outputPath = @"C:\Docs\TemplateCopy.docx";
        doc.Save(outputPath);
    }
}
