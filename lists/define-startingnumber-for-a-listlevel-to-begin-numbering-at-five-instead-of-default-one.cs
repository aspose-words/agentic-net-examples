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

        // Create a numbered list using the default number template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Set the starting number of the first level to 5.
        list.ListLevels[0].StartAt = 5;

        // Apply the list to the builder.
        builder.ListFormat.List = list;

        // Write three list items; they will be numbered 5, 6, 7.
        builder.Writeln("First item");
        builder.Writeln("Second item");
        builder.Writeln("Third item");

        // Remove list formatting after the list is complete.
        builder.ListFormat.RemoveNumbers();

        // Save the document to a file.
        doc.Save("Result.docx");
    }
}
