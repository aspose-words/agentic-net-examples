using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a numbered list and add three items.
        builder.ListFormat.ApplyNumberDefault();
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");
        builder.Writeln("Item 3");

        // Stop list formatting for subsequent paragraphs.
        builder.ListFormat.RemoveNumbers();

        // Add a normal paragraph after the list.
        builder.Writeln("This paragraph is not part of the list.");

        // Ensure any existing list formatting is removed from all paragraphs.
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            if (para.ListFormat.IsListItem)
                para.ListFormat.RemoveNumbers();
        }

        // Save the document.
        string outputPath = "RemoveNumbersExample.docx";
        doc.Save(outputPath);
    }
}
