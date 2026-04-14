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

        // Stop list formatting for any following paragraphs.
        builder.ListFormat.RemoveNumbers();

        // Add a regular paragraph after the list.
        builder.Writeln("This paragraph is not part of the list.");

        // Ensure all existing list paragraphs have their numbering removed.
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            if (para.ListFormat.IsListItem)
                para.ListFormat.RemoveNumbers();
        }

        // Save the document to an output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "RemoveNumbersExample.docx");
        doc.Save(outputPath);
    }
}
