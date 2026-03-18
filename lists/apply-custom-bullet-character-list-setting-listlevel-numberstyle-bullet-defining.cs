using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Lists;

class CustomBulletExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a DocumentBuilder to insert content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a new list based on the default bullet template.
        List list = doc.Lists.Add(ListTemplate.BulletDefault);

        // Choose a custom bullet character (e.g., a black star).
        string bulletChar = "\u2605"; // Unicode star character

        // Configure the first level of the list:
        // - Set the number style to Bullet.
        // - Set the NumberFormat to the custom bullet character.
        // - Optionally adjust font to ensure the character displays correctly.
        ListLevel level = list.ListLevels[0];
        level.NumberStyle = NumberStyle.Bullet;
        level.NumberFormat = bulletChar;
        level.Font.Name = "Arial";
        level.Font.Size = 12;
        level.Font.Color = Color.Black;

        // Apply the list to subsequent paragraphs.
        builder.ListFormat.List = list;

        // Add some list items that will use the custom bullet.
        builder.Writeln("First item with custom bullet");
        builder.Writeln("Second item with custom bullet");
        builder.Writeln("Third item with custom bullet");

        // End the list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the document to disk.
        doc.Save("CustomBulletList.docx");
    }
}
