using System;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -------------------------------------------------
        // Create a bullet list and assign it to paragraphs.
        // -------------------------------------------------
        List bulletList = doc.Lists.Add(ListTemplate.BulletDefault);

        // Apply the existing list to the builder.
        builder.ListFormat.List = bulletList;
        builder.Writeln("Bullet item 1");
        builder.Writeln("Bullet item 2");

        // Finish the bullet list.
        builder.ListFormat.RemoveNumbers();

        // -------------------------------------------------
        // Create a numbered list and assign it to paragraphs.
        // -------------------------------------------------
        List numberList = doc.Lists.Add(ListTemplate.NumberDefault);

        // Apply the existing numbered list.
        builder.ListFormat.List = numberList;
        builder.Writeln("Numbered item 1");
        builder.Writeln("Numbered item 2");

        // Finish the numbered list.
        builder.ListFormat.RemoveNumbers();

        // -------------------------------------------------
        // Validation: count paragraphs that are part of a list.
        // -------------------------------------------------
        int listParagraphCount = 0;
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            if (para.ListFormat.IsListItem)
                listParagraphCount++;
        }

        Console.WriteLine($"Document contains {listParagraphCount} list paragraphs.");

        // Save the document to a file.
        doc.Save("Output.docx");
    }
}
