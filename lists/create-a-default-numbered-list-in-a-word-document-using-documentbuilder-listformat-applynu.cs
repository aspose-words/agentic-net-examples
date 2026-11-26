using System;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a default numbered list (1., 2., 3., ...).
        List defaultNumberedList = doc.Lists.Add(ListTemplate.NumberDefault);

        // Assign the list to the builder so that subsequent paragraphs belong to this list.
        builder.ListFormat.List = defaultNumberedList;

        // Write three list items, applying the default numbering to each paragraph.
        for (int i = 1; i <= 3; i++)
        {
            builder.ListFormat.ApplyNumberDefault();   // Apply default numbering style.
            builder.Writeln($"Item {i}");
        }

        // Remove list formatting after the list is complete.
        builder.ListFormat.RemoveNumbers();

        // Optional validation: count paragraphs that are part of a list.
        int listParagraphCount = 0;
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            if (para.ListFormat.IsListItem)
                listParagraphCount++;
        }
        Console.WriteLine($"Number of list paragraphs: {listParagraphCount}");

        // Save the document to a file.
        doc.Save("DefaultNumberedList.docx");
    }
}
