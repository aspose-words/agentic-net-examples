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

        // Create a DocumentBuilder to insert content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a list based on the default numbered template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Adjust the indentation of the first list level.
        // Aspose.Words does not have a direct Indentation property.
        // Instead, set NumberPosition (position of the number/bullet) and TextPosition (position of the text).
        // Setting NumberPosition to -36 points moves the number left, and TextPosition to 36 points
        // creates a total left indent of 36 points for the list item text.
        ListLevel level = list.ListLevels[0];
        level.NumberPosition = -36; // Move the number/bullet left.
        level.TextPosition = 36;    // Position the text 36 points from the left margin.

        // Apply the list to the builder and add some items.
        builder.ListFormat.List = list;
        builder.Writeln("First item with custom indentation.");
        builder.Writeln("Second item with custom indentation.");

        // Remove list formatting from subsequent paragraphs.
        builder.ListFormat.RemoveNumbers();

        // Prepare output directory.
        string artifactsDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(artifactsDir);

        // Save the document.
        doc.Save(Path.Combine(artifactsDir, "ListIndentation.docx"));
    }
}
