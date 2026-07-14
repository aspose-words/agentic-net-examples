using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Use DocumentBuilder to add sample lists to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a numbered list.
        builder.ListFormat.ApplyNumberDefault();
        builder.Writeln("Numbered item 1");
        builder.Writeln("Numbered item 2");
        builder.ListFormat.RemoveNumbers();

        // Add a bulleted list.
        builder.ListFormat.ApplyBulletDefault();
        builder.Writeln("Bullet item A");
        builder.Writeln("Bullet item B");
        builder.ListFormat.RemoveNumbers();

        // Iterate over every list in the document.
        foreach (List list in doc.Lists)
        {
            // Apply the same formatting to each level of the current list.
            foreach (ListLevel level in list.ListLevels)
            {
                level.Font.Name = "Arial";
                level.Font.Color = Color.Green;
                level.Font.Bold = true;
            }
        }

        // Save the modified document to the file system.
        doc.Save("UniformListStyle.docx");
    }
}
