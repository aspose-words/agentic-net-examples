using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a DocumentBuilder which will be used to insert content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a new list based on the default bullet template.
        List list = doc.Lists.Add(ListTemplate.BulletDefault);

        // Access the first (0‑based) level of the list.
        ListLevel level = list.ListLevels[0];

        // Set the level to use a bullet style.
        level.NumberStyle = NumberStyle.Bullet;

        // Define a custom bullet character. Here we use the Unicode black circle (U+2022).
        level.NumberFormat = "\u2022";

        // Optionally change the font used for the bullet.
        level.Font.Name = "Wingdings";
        level.Font.Size = 12;

        // Apply the customized list to the builder and add some items.
        builder.ListFormat.List = list;
        builder.Writeln("Custom bullet item 1");
        builder.Writeln("Custom bullet item 2");
        builder.Writeln("Custom bullet item 3");

        // End the list.
        builder.ListFormat.RemoveNumbers();

        // Save the document to disk.
        doc.Save("CustomBulletList.docx");
    }
}
