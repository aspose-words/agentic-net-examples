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

        // Add a list based on the default bullet template.
        List list = doc.Lists.Add(ListTemplate.BulletDefault);

        // Customize the first level to use a dash ("-") as the bullet character.
        ListLevel level = list.ListLevels[0];
        level.NumberStyle = NumberStyle.Bullet;
        level.NumberFormat = "-";
        level.TrailingCharacter = ListTrailingCharacter.Space;

        // Apply the customized list to the builder.
        builder.ListFormat.List = list;

        // Add three list items.
        builder.Writeln("First item");
        builder.Writeln("Second item");
        builder.Writeln("Third item");

        // End the list.
        builder.ListFormat.RemoveNumbers();

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "BulletedList.docx");

        // Save the document.
        doc.Save(outputPath);
    }
}
