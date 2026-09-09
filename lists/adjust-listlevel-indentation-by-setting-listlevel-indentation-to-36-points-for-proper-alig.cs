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
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a list based on the default numbered template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Adjust the indentation of the first two list levels to 36 points.
        // The ListLevel class does not have an Indentation property.
        // Use NumberPosition (position of the bullet/number) and TextPosition (position of the text)
        // to achieve the desired left indent.
        ListLevel level0 = list.ListLevels[0];
        level0.NumberPosition = 36; // points
        level0.TextPosition = 36;   // points

        ListLevel level1 = list.ListLevels[1];
        level1.NumberPosition = 36; // points
        level1.TextPosition = 36;   // points

        // Apply the list to some paragraphs.
        builder.ListFormat.List = list;
        builder.Writeln("Item 1");
        builder.ListFormat.ListIndent();
        builder.Writeln("Subitem 1");
        builder.ListFormat.ListOutdent();
        builder.Writeln("Item 2");
        builder.ListFormat.RemoveNumbers();

        // Ensure the output directory exists.
        string artifactsDir = Path.Combine(Environment.CurrentDirectory, "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Save the document.
        doc.Save(Path.Combine(artifactsDir, "AdjustedListIndentation.docx"));
    }
}
