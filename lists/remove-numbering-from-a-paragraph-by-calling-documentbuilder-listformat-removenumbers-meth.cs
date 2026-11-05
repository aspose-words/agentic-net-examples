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

        // Create a numbered list using a built‑in list template.
        List numberedList = doc.Lists.Add(ListTemplate.NumberDefault);

        // Apply the list to the builder so that subsequent paragraphs become list items.
        builder.ListFormat.List = numberedList;

        // Write several items that will be numbered.
        builder.Writeln("First item");
        builder.Writeln("Second item");
        builder.Writeln("Third item");

        // Stop list formatting. This removes numbering from the next paragraph.
        builder.ListFormat.RemoveNumbers();

        // Write a normal paragraph that is not part of the list.
        builder.Writeln("This paragraph is not part of the list.");

        // Save the document to disk.
        doc.Save("Output.docx");
    }
}
