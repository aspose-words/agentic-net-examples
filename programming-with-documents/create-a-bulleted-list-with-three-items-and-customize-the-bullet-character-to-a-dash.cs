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

        // Add a new list based on a bullet template.
        List list = doc.Lists.Add(ListTemplate.BulletDefault);
        ListLevel level = list.ListLevels[0];

        // Customize the bullet character to a dash.
        level.NumberFormat = "-";          // Set dash as the bullet character.
        level.Font.Name = "Arial";         // Font for the bullet.
        level.Font.Size = 12;              // Size for the bullet.

        // Apply the list to the builder.
        builder.ListFormat.List = list;

        // Write three list items.
        builder.Writeln("First item");
        builder.Writeln("Second item");
        builder.Writeln("Third item");

        // End the list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the document to a file.
        string outputPath = "BulletedList.docx";
        doc.Save(outputPath);
    }
}
