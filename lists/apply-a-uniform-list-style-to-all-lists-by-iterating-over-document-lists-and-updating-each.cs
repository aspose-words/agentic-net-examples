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

        // Use DocumentBuilder to add a couple of sample lists.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First list – numbered.
        builder.ListFormat.ApplyNumberDefault();
        builder.Writeln("First numbered item");
        builder.Writeln("Second numbered item");
        builder.ListFormat.RemoveNumbers();

        // Second list – bulleted.
        builder.ListFormat.ApplyBulletDefault();
        builder.Writeln("First bullet item");
        builder.Writeln("Second bullet item");
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

        // Save the document to the local file system.
        string outputPath = "Lists_UniformStyle.docx";
        doc.Save(outputPath);
    }
}
