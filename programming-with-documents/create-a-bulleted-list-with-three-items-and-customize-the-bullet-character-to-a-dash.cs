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

        // Start a bullet list using the default bullet style.
        builder.ListFormat.ApplyBulletDefault();

        // Get the current list level and customize its bullet character to a dash.
        ListLevel level = builder.ListFormat.ListLevel;
        // The bullet style is already set by ApplyBulletDefault().
        // Change the bullet character to a dash.
        level.NumberFormat = "-";

        // Add three list items.
        builder.Writeln("First item");
        builder.Writeln("Second item");
        builder.Writeln("Third item");

        // End the list.
        builder.ListFormat.RemoveNumbers();

        // Save the document to a file.
        string outputPath = "BulletedList.docx";
        doc.Save(outputPath);
    }
}
