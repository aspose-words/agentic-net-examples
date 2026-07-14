using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new empty document and a DocumentBuilder to edit it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a default numbered list and add three items.
        builder.ListFormat.ApplyNumberDefault();
        builder.Writeln("Numbered item 1");
        builder.Writeln("Numbered item 2");
        builder.Writeln("Numbered item 3");

        // Remove list formatting (numbers/bullets) from every paragraph in the document.
        foreach (Paragraph paragraph in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            paragraph.ListFormat.RemoveNumbers();
        }

        // Save the resulting document to disk.
        doc.Save("RemoveNumbers.docx");
    }
}
