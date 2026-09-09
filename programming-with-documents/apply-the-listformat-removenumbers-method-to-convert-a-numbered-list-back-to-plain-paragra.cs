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

        // Start a default numbered list and add three items.
        builder.ListFormat.ApplyNumberDefault();
        builder.Writeln("Numbered list item 1");
        builder.Writeln("Numbered list item 2");
        builder.Writeln("Numbered list item 3");

        // Convert each list item back to a plain paragraph by removing its list formatting.
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        foreach (Paragraph paragraph in paragraphs)
        {
            paragraph.ListFormat.RemoveNumbers();
        }

        // Save the resulting document.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "ListRemoved.docx");
        doc.Save(outputPath);
    }
}
