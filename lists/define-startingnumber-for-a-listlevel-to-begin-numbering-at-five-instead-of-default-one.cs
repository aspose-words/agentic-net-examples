using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a list based on the default numbered template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Set the starting number of the first list level to 5.
        list.ListLevels[0].StartAt = 5;

        // Use DocumentBuilder to add paragraphs that use the custom list.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("List with custom starting number:");
        builder.ListFormat.List = list;
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");
        builder.Writeln("Item 3");
        builder.ListFormat.RemoveNumbers();

        // Define the output path (current directory).
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ListStartAt.docx");

        // Save the document.
        doc.Save(outputPath);
    }
}
