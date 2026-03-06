using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables; // For NodeType enum

class Program
{
    static void Main()
    {
        // Path to the DOT template file.
        string templatePath = @"C:\Docs\Template.dot";

        // Load the DOT document using the Document constructor (load rule).
        Document doc = new Document(templatePath);

        // Retrieve all Paragraph nodes in the document (including those in headers/footers).
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true)
                           .Cast<Paragraph>()
                           .Select(p => p.GetText().Trim())
                           .Where(text => !string.IsNullOrEmpty(text));

        // Output each paragraph text to the console.
        foreach (var text in paragraphs)
        {
            Console.WriteLine(text);
        }

        // (Optional) Save the document after processing, if needed.
        // string outputPath = @"C:\Docs\Processed.docx";
        // doc.Save(outputPath);
    }
}
