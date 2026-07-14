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
        builder.Writeln("Numbered list item 1");
        builder.Writeln("Numbered list item 2");
        builder.Writeln("Numbered list item 3");

        // Remove list formatting from all paragraphs in the document.
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            if (para.ListFormat.IsListItem)
                para.ListFormat.RemoveNumbers();
        }

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Artifacts");
        Directory.CreateDirectory(outputDir);

        // Save the resulting document.
        string outputPath = Path.Combine(outputDir, "ListRemoved.docx");
        doc.Save(outputPath);
    }
}
