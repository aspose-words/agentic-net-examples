using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a numbered list with three items.
        builder.ListFormat.ApplyNumberDefault();
        builder.Writeln("First item");
        builder.Writeln("Second item");
        builder.Writeln("Third item");
        // Stop applying list formatting for subsequent paragraphs.
        builder.ListFormat.RemoveNumbers();

        // Convert the numbered list back to plain paragraphs.
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            if (para.IsListItem)
            {
                para.ListFormat.RemoveNumbers();
            }
        }

        // Save the resulting document.
        const string outputPath = "Output.docx";
        doc.Save(outputPath);
    }
}
