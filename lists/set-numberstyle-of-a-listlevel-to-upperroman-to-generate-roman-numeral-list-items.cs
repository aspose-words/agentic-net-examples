using System;
using System.Drawing;
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

        // Set the first list level to use uppercase Roman numerals.
        ListLevel level = list.ListLevels[0];
        level.NumberStyle = NumberStyle.UppercaseRoman;
        level.NumberFormat = "\x0000."; // Inserts the Roman numeral followed by a period.
        level.Font.Size = 12;
        level.Font.Color = Color.Black;

        // Apply the list to a series of paragraphs.
        builder.ListFormat.List = list;
        for (int i = 0; i < 5; i++)
        {
            builder.Writeln($"Item {i + 1}");
        }

        // End the list.
        builder.ListFormat.RemoveNumbers();

        // Save the document to the current directory.
        doc.Save("ListRoman.docx");
    }
}
