using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a numbered list and add three items.
        builder.ListFormat.ApplyNumberDefault();
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");
        builder.Writeln("Item 3");

        // End the list so further paragraphs are not automatically part of it.
        builder.ListFormat.RemoveNumbers();

        // Add a regular paragraph after the list.
        builder.Writeln("Normal paragraph after list.");

        // Retrieve all paragraphs in the document.
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

        // Remove list formatting from the second and third list items while preserving indentation.
        // Paragraph indices: 0 = "Item 1", 1 = "Item 2", 2 = "Item 3", 3 = normal paragraph.
        for (int i = 1; i <= 2; i++)
        {
            Paragraph para = (Paragraph)paragraphs[i];
            if (para.ListFormat.IsListItem)
            {
                // This call removes the number/bullet and sets the list level to zero,
                // but the left indent of the paragraph remains unchanged.
                para.ListFormat.RemoveNumbers();
            }
        }

        // Save the resulting document.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "RemoveListFormatting.docx");
        doc.Save(outputPath);
    }
}
