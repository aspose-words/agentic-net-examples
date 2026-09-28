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

        // Add a numbered list to the document's list collection.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // First paragraph – assign the list to this paragraph.
        builder.Writeln("First item");
        Paragraph para1 = builder.CurrentParagraph; // Get the paragraph just created.
        para1.ListFormat.List = list;               // Apply the list style.
        para1.ListFormat.ListLevelNumber = 0;       // Use the first level of the list.

        // Second paragraph – also part of the same list.
        builder.Writeln("Second item");
        Paragraph para2 = builder.CurrentParagraph;
        para2.ListFormat.List = list;
        para2.ListFormat.ListLevelNumber = 0;

        // Third paragraph – a sub‑item (level 1).
        builder.Writeln("Sub‑item");
        Paragraph para3 = builder.CurrentParagraph;
        para3.ListFormat.List = list;
        para3.ListFormat.ListLevelNumber = 1;

        // Save the document to disk.
        doc.Save("ListExample.docx");
    }
}
