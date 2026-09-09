using System;
using Aspose.Words;
using Aspose.Words.Lists;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Create a new list based on the default bullet template.
        List bulletList = doc.Lists.Add(ListTemplate.BulletDefault);

        // Access the first (top) level of the list.
        ListLevel level = bulletList.ListLevels[0];

        // Set the level to use a bullet style.
        level.NumberStyle = NumberStyle.Bullet;

        // Define a custom bullet character (★ - black star).
        level.NumberFormat = "\u2605";

        // Optional: customize the appearance of the bullet.
        level.Font.Name = "Arial";
        level.Font.Size = 12;
        level.Font.Color = Color.DarkBlue;

        // Use a DocumentBuilder to add paragraphs that will use the custom bullet list.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.ListFormat.List = bulletList;

        builder.Writeln("First item with a custom bullet.");
        builder.Writeln("Second item with a custom bullet.");
        builder.Writeln("Third item with a custom bullet.");

        // End the list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the document to a file.
        doc.Save("CustomBulletList.docx");
    }
}
