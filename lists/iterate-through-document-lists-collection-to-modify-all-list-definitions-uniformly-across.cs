using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a numbered list.
        builder.Writeln("Numbered List:");
        List numberedList = doc.Lists.Add(ListTemplate.NumberDefault);
        builder.ListFormat.List = numberedList;
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");
        builder.ListFormat.RemoveNumbers();

        // Add a bulleted list.
        builder.Writeln("Bulleted List:");
        List bulletedList = doc.Lists.Add(ListTemplate.BulletDefault);
        builder.ListFormat.List = bulletedList;
        builder.Writeln("Bullet 1");
        builder.Writeln("Bullet 2");
        builder.ListFormat.RemoveNumbers();

        // Iterate through all lists in the document and modify them uniformly.
        foreach (List list in doc.Lists)
        {
            // Restart numbering at each section.
            list.IsRestartAtEachSection = true;

            // Change the font color of the first level of each list to green.
            if (list.ListLevels.Count > 0)
            {
                list.ListLevels[0].Font.Color = Color.Green;
            }
        }

        // Save the modified document.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "ModifiedLists.docx");
        doc.Save(outputPath);
    }
}
