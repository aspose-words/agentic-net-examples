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

        // Iterate over every list in the document and set a uniform style for each level.
        foreach (List list in doc.Lists)
        {
            foreach (ListLevel level in list.ListLevels)
            {
                level.Font.Name = "Arial";
                level.Font.Size = 12;
                level.Font.Color = Color.Black;
                level.Font.Bold = false;
            }
        }

        // Save the resulting document to the file system.
        doc.Save("UniformLists.docx");
    }
}
