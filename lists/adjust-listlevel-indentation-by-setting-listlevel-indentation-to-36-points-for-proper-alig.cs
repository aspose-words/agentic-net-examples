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

        // Create a numbered list using a built‑in template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Adjust the indentation of the first list level to 36 points.
        // Aspose.Words does not expose an "Indentation" property directly.
        // Instead, set NumberPosition (position of the number) and TextPosition (position of the list text).
        ListLevel level0 = list.ListLevels[0];
        level0.NumberPosition = 36; // distance from left margin to the number
        level0.TextPosition = 36;   // distance from left margin to the list text

        // Apply the list to the builder.
        builder.ListFormat.List = list;
        builder.ListFormat.ListLevelNumber = 0; // first level

        // Write list items.
        builder.Writeln("First item");
        builder.Writeln("Second item");
        builder.Writeln("Third item");

        // Remove list formatting after the list is complete.
        builder.ListFormat.RemoveNumbers();

        // Save the document to a file.
        doc.Save("ListIndentationExample.docx");

        // Indicate completion.
        Console.WriteLine("Document created with list indentation set to 36 points.");
    }
}
