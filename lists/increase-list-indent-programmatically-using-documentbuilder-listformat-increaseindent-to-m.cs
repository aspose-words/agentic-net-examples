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

        // Create a numbered list and assign it to the builder.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);
        builder.ListFormat.List = list;

        // First level list item.
        builder.ListFormat.ListLevelNumber = 0; // top level
        builder.Writeln("Item 1");

        // Increase indent to move to a deeper level within the same list.
        builder.ListFormat.ListLevelNumber = 1; // second level
        builder.Writeln("Item 1.1");

        // Decrease indent back to the original level.
        builder.ListFormat.ListLevelNumber = 0; // back to top level
        builder.Writeln("Item 2");

        // Finish the list and remove list formatting from subsequent paragraphs.
        builder.ListFormat.RemoveNumbers();

        // Validation: count paragraphs that are part of the list.
        int listParagraphCount = 0;
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            if (para.ListFormat.IsListItem)
                listParagraphCount++;
        }

        Console.WriteLine($"The document contains {listParagraphCount} list paragraphs.");

        // Save the document to a file.
        doc.Save("ListIndent.docx");
    }
}
