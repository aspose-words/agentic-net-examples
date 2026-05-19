using System;
using Aspose.Words;
using Aspose.Words.Lists;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a numbered list based on the default template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Set the first list level to start numbering at 5 instead of the default 1.
        list.ListLevels[0].StartAt = 5;

        // Write a heading and apply the custom list to a few items.
        builder.Writeln("List that starts at 5:");
        builder.ListFormat.List = list;
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");
        builder.Writeln("Item 3");
        builder.ListFormat.RemoveNumbers();

        // Save the document to the current directory.
        doc.Save("ListStartAtFive.docx");
    }
}
