using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Add a numbered list based on the default template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Set the first list level to start numbering at 5.
        list.ListLevels[0].StartAt = 5;

        // Use DocumentBuilder to write paragraphs that use the list.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("List that starts at 5:");
        builder.ListFormat.List = list;
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");
        builder.Writeln("Item 3");
        builder.ListFormat.RemoveNumbers();

        // Save the document to the file system.
        doc.Save("ListStartAtFive.docx");
    }
}
