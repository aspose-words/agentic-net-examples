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

        // Add a numbered list to the document.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Adjust the indentation of the first list level to 36 points.
        // In Aspose.Words the indentation is controlled by NumberPosition (position of the number)
        // and TextPosition (position of the list text). Both are set in points.
        ListLevel level = list.ListLevels[0];
        level.NumberPosition = 36f; // position of the number
        level.TextPosition = 36f;   // position of the text after the number

        // Apply the list to subsequent paragraphs.
        builder.ListFormat.List = list;
        builder.Writeln("First item");
        builder.Writeln("Second item");
        builder.Writeln("Third item");

        // Remove list formatting for following text.
        builder.ListFormat.RemoveNumbers();

        // Save the document to disk.
        doc.Save("ListIndentation.docx");
    }
}
