using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Use DocumentBuilder to add a numbered list.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.ListFormat.ApplyNumberDefault(); // start default numbered list

        // Add three list items.
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");
        builder.Writeln("Item 3");

        // Retrieve all paragraphs in the document.
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

        // Remove list formatting from the second paragraph (index 1) while keeping its indentation.
        if (paragraphs.Count > 1 && paragraphs[1] is Paragraph paraToModify)
        {
            paraToModify.ListFormat.RemoveNumbers();
        }

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Lists_RemoveNumbers.docx");
        doc.Save(outputPath);
    }
}
