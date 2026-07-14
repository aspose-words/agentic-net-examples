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
        List customList = doc.Lists.Add(ListTemplate.BulletDefault);

        // Access the first (level 0) list level to customize its bullet.
        ListLevel level = customList.ListLevels[0];
        // Set the list level to use a bullet style.
        level.NumberStyle = NumberStyle.Bullet;
        // Define a custom bullet character. Here we use the check mark (✓) as an example.
        level.NumberFormat = "\u2713";
        // Optional: set the font that contains the custom character.
        level.Font.Name = "Wingdings";
        level.Font.Size = 12;
        level.Font.Color = Color.DarkGreen;

        // Use DocumentBuilder to add paragraphs that will use the custom list.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.ListFormat.List = customList; // Apply the list to subsequent paragraphs.

        builder.Writeln("First custom bullet item");
        builder.Writeln("Second custom bullet item");
        builder.Writeln("Third custom bullet item");

        // Remove list formatting from further paragraphs.
        builder.ListFormat.RemoveNumbers();
        builder.Writeln("This paragraph is not part of the list.");

        // Save the document to the current directory.
        doc.Save("CustomBulletList.docx");
    }
}
