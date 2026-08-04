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

        // Initialize a DocumentBuilder which will be used to insert content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a new list based on the default bullet template.
        List bulletList = doc.Lists.Add(ListTemplate.BulletDefault);

        // Access the first (top) level of the list.
        ListLevel level = bulletList.ListLevels[0];

        // Set the list level to use a bullet style.
        level.NumberStyle = NumberStyle.Bullet;

        // Define a custom bullet character (black circle in this case).
        // The NumberFormat property holds the character that will be displayed as the bullet.
        level.NumberFormat = "\u25CF"; // Unicode character: ●

        // Optionally, set the font that contains the bullet glyph.
        level.Font.Name = "Symbol";

        // Apply the custom list to the builder.
        builder.ListFormat.List = bulletList;

        // Add some list items.
        builder.Writeln("First custom bullet item");
        builder.Writeln("Second custom bullet item");
        builder.Writeln("Third custom bullet item");

        // End the list.
        builder.ListFormat.RemoveNumbers();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CustomBulletList.docx");
        doc.Save(outputPath);
    }
}
