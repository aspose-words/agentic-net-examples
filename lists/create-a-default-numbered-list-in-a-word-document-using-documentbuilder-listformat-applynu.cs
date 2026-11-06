using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a default numbered list (1., 2., 3., ...).
        List numberedList = doc.Lists.Add(ListTemplate.NumberDefault);

        // Apply the list to the builder so subsequent paragraphs become list items.
        builder.ListFormat.List = numberedList;

        // Write the list items.
        builder.Writeln("First item");
        builder.Writeln("Second item");
        builder.Writeln("Third item");

        // Remove list formatting after the list is finished.
        builder.ListFormat.RemoveNumbers();

        // Save the document to disk.
        doc.Save("NumberedList.docx");
    }
}
