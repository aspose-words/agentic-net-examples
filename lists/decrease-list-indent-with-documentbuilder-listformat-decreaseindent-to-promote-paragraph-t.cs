using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a numbered list using a built‑in list template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Apply the list to the builder.
        builder.ListFormat.List = list;

        // Write a top‑level list item.
        builder.ListFormat.ListLevelNumber = 0; // top level
        builder.Writeln("Item 1");

        // Increase indent to create a sub‑item (level 1).
        builder.ListFormat.ListLevelNumber = 1; // now level 1
        builder.Writeln("Item 1.1");

        // Decrease indent to promote back to the top level.
        builder.ListFormat.ListLevelNumber = 0; // back to level 0
        builder.Writeln("Item 2");

        // Finish the list and remove list formatting from the builder.
        builder.ListFormat.RemoveNumbers();

        // Validation: count paragraphs that are part of a list.
        int listParagraphCount = 0;
        foreach (Paragraph para in doc.FirstSection.Body.Paragraphs)
        {
            if (para.ListFormat.IsListItem)
                listParagraphCount++;
        }

        Console.WriteLine($"Document contains {listParagraphCount} list paragraphs.");

        // Save the document to a file.
        doc.Save("Output.docx");
    }
}
