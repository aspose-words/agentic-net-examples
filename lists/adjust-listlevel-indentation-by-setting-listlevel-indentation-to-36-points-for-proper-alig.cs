using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Create a DocumentBuilder which will be used to insert content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a new list based on the default numbered template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Adjust the indentation of the first list level to 36 points.
        // Since ListLevel does not have an Indentation property, we set the
        // NumberPosition (position of the number) and TextPosition (position of the text)
        // to achieve the desired left indent.
        ListLevel level = list.ListLevels[0];
        level.NumberPosition = 0;      // Number starts at the left margin.
        level.TextPosition = 36;       // Text starts 36 points to the right.
        level.TabPosition = 36;        // Tab stop aligns with the text position.

        // Apply the list to the builder and add a few items.
        builder.ListFormat.List = list;
        builder.Writeln("First item with custom indentation");
        builder.Writeln("Second item with custom indentation");
        builder.ListFormat.RemoveNumbers();

        // Save the document to the output file.
        string outputPath = System.IO.Path.Combine(Environment.CurrentDirectory, "ListIndentation.docx");
        doc.Save(outputPath);
    }
}
