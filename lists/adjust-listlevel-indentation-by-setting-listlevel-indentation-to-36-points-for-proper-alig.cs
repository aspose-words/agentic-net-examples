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

        // Add a multi‑level list based on the default numbered template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Adjust the indentation of the first list level.
        // The ListLevel class does not expose an "Indentation" property.
        // Instead, we set the positions that control the left indent:
        //   NumberPosition – position of the number/bullet (negative value moves it left of the text).
        //   TextPosition   – position where the list text starts.
        //   TabPosition    – tab stop after the number/bullet.
        // Setting these to achieve a 36‑point left indent.
        ListLevel level = list.ListLevels[0];
        level.NumberPosition = -36; // number placed 36 points to the left of the text start
        level.TextPosition = 36;    // text starts 36 points from the left margin
        level.TabPosition = 36;     // tab stop aligns with the text start

        // Apply the list to a couple of paragraphs.
        builder.ListFormat.List = list;
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");
        builder.ListFormat.RemoveNumbers();

        // Save the document to the output file.
        doc.Save("AdjustedListIndentation.docx");
    }
}
