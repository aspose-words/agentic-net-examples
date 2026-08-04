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

        // Build a few sample lists so that we have something to modify.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Numbered list.
        builder.ListFormat.ApplyNumberDefault();
        builder.Writeln("Numbered item 1");
        builder.Writeln("Numbered item 2");
        builder.ListFormat.RemoveNumbers();

        // Bulleted list.
        builder.ListFormat.ApplyBulletDefault();
        builder.Writeln("Bulleted item 1");
        builder.Writeln("Bulleted item 2");
        builder.ListFormat.RemoveNumbers();

        // Iterate over all lists in the document.
        foreach (List list in doc.Lists)
        {
            // Apply the same formatting to every level of the current list.
            foreach (ListLevel level in list.ListLevels)
            {
                level.Font.Name = "Arial";
                level.Font.Color = Color.Green;
                level.Font.Bold = true;
            }
        }

        // Save the resulting document.
        string outputPath = "UniformListStyle.docx";
        doc.Save(outputPath);
    }
}
