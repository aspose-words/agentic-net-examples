using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a new list based on the default bullet template.
        List bulletList = doc.Lists.Add(ListTemplate.BulletDefault);

        // Customize the first level of the list to use a dash ("-") as the bullet character.
        // The NumberStyle must be set to Bullet and the NumberFormat to the desired symbol.
        bulletList.ListLevels[0].NumberStyle = NumberStyle.Bullet;
        bulletList.ListLevels[0].NumberFormat = "-";

        // Apply the customized list to the builder.
        builder.ListFormat.List = bulletList;
        builder.ListFormat.ListLevelNumber = 0; // Ensure we are on the first level.

        // Add three list items.
        builder.Writeln("First item");
        builder.Writeln("Second item");
        builder.Writeln("Third item");

        // End the list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the document to a file.
        doc.Save("BulletedList.docx");
    }
}
