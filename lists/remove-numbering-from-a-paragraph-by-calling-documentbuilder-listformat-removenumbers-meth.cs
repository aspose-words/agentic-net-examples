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
        List numberedList = doc.Lists.Add(ListTemplate.NumberDefault);
        builder.ListFormat.List = numberedList;

        // Add list items.
        builder.Writeln("First item");
        builder.Writeln("Second item");

        // Stop list formatting – subsequent paragraphs will not be numbered.
        builder.ListFormat.RemoveNumbers();

        // Add a regular paragraph that is not part of the list.
        builder.Writeln("This paragraph is not part of the list.");

        // Save the document to disk.
        doc.Save("Output.docx");
    }
}
