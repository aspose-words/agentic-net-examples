using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add a new list based on the default bullet template.
        List list = doc.Lists.Add(ListTemplate.BulletDefault);

        // Access the first (top) level of the list.
        ListLevel level = list.ListLevels[0];

        // Set the list level to use a bullet style.
        level.NumberStyle = NumberStyle.Bullet;

        // Choose a font that contains the desired bullet character.
        level.Font.Name = "Wingdings";
        level.Font.Size = 12;

        // Define a custom bullet character (e.g., a check mark).
        // The character is specified using its Unicode escape sequence.
        level.NumberFormat = "\u2713"; // ✔

        // Use DocumentBuilder to add paragraphs that will be formatted as list items.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.ListFormat.List = list; // Apply the custom list to the builder.
        builder.Writeln("First custom bullet item");
        builder.Writeln("Second custom bullet item");
        builder.Writeln("Third custom bullet item");

        // End the list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the document to the current directory.
        string outputPath = System.IO.Path.Combine(Environment.CurrentDirectory, "CustomBulletList.docx");
        doc.Save(outputPath);
    }
}
