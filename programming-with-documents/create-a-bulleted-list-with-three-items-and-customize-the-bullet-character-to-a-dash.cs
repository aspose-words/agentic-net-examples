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

        // Create a new list based on the default bullet template.
        List list = doc.Lists.Add(ListTemplate.BulletDefault);

        // Customize the first level of the list to use a dash as the bullet character.
        ListLevel level = list.ListLevels[0];
        // The default NumberStyle for a bullet list is already set, so we only need to change the format.
        level.NumberFormat = "-";
        level.Font.Name = "Arial";

        // Apply the list to the builder.
        builder.ListFormat.List = list;

        // Add three list items.
        builder.Writeln("First item");
        builder.Writeln("Second item");
        builder.Writeln("Third item");

        // End the list formatting.
        builder.ListFormat.RemoveNumbers();

        // Define output file path.
        string outputPath = "BulletedList.docx";

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved successfully to '{outputPath}'.");
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
