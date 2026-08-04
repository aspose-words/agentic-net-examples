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
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a single‑level bulleted list based on the default bullet template.
        List bulletList = doc.Lists.AddSingleLevelList(ListTemplate.BulletDefault);

        // Customize the first (and only) level to use a dash ("-") as the bullet character.
        ListLevel level = bulletList.ListLevels[0];
        level.NumberStyle = NumberStyle.Bullet;   // Ensure the level is treated as a bullet.
        level.NumberFormat = "-";                 // The dash character will be used as the bullet.
        level.TrailingCharacter = ListTrailingCharacter.Space; // Add a space after the dash.

        // Apply the customized list to the builder.
        builder.ListFormat.List = bulletList;
        builder.ListFormat.ListLevelNumber = 0; // Use the first level.

        // Write three list items.
        builder.Writeln("First item");
        builder.Writeln("Second item");
        builder.Writeln("Third item");

        // End the list formatting.
        builder.ListFormat.RemoveNumbers();

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document.
        string outputPath = Path.Combine(outputDir, "BulletedList.docx");
        doc.Save(outputPath);
    }
}
