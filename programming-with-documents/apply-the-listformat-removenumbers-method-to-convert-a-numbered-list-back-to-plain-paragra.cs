using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create an output folder for the generated document.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(artifactsDir);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a numbered list and add three items.
        builder.ListFormat.ApplyNumberDefault();
        builder.Writeln("Numbered list item 1");
        builder.Writeln("Numbered list item 2");
        builder.Writeln("Numbered list item 3");

        // Convert the numbered list back to plain paragraphs.
        // Iterate over all paragraphs and remove list formatting where present.
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            if (para.ListFormat.IsListItem)
                para.ListFormat.RemoveNumbers();
        }

        // Save the document.
        string outputPath = Path.Combine(artifactsDir, "ListRemoved.docx");
        doc.Save(outputPath);
    }
}
