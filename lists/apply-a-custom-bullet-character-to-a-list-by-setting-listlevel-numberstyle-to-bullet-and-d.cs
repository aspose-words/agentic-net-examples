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

        // Add a bulleted list based on the default bullet template.
        List list = doc.Lists.Add(ListTemplate.BulletDefault);

        // Customize the first level of the list.
        ListLevel level = list.ListLevels[0];
        level.NumberStyle = NumberStyle.Bullet;          // Use bullet numbering.
        level.NumberFormat = "\xf0af";                    // Custom bullet character (star in Wingdings).
        level.Font.Name = "Wingdings";                    // Font that contains the custom glyph.
        level.Font.Size = 12;                             // Optional: set bullet size.
        level.TrailingCharacter = ListTrailingCharacter.Space; // Space after the bullet.
        level.NumberPosition = 0;                         // Position of the bullet.

        // Apply the customized list to paragraphs.
        builder.ListFormat.List = list;
        builder.Writeln("First custom bullet item");
        builder.Writeln("Second custom bullet item");
        builder.Writeln("Third custom bullet item");

        // End the list.
        builder.ListFormat.RemoveNumbers();

        // Save the document to a file.
        doc.Save("CustomBulletList.docx");
    }
}
