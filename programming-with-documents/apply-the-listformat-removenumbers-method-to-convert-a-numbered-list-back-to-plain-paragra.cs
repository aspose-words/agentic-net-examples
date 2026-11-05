using System;
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
        builder.Writeln("First item");
        builder.Writeln("Second item");
        builder.Writeln("Third item");
        // Stop list formatting for any following text.
        builder.ListFormat.RemoveNumbers();

        // Convert the numbered list items back to plain paragraphs.
        foreach (Paragraph paragraph in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            if (paragraph.IsListItem)
            {
                paragraph.ListFormat.RemoveNumbers();
            }
        }

        // Save the resulting document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);
    }
}
