using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a numbered list.
        List numberedList = doc.Lists.Add(ListTemplate.NumberDefault);
        builder.ListFormat.List = numberedList;
        for (int i = 1; i <= 3; i++)
            builder.Writeln($"Numbered item {i}");
        builder.ListFormat.RemoveNumbers();

        // Add a bulleted list.
        List bulletList = doc.Lists.Add(ListTemplate.BulletDefault);
        builder.ListFormat.List = bulletList;
        for (int i = 1; i <= 3; i++)
            builder.Writeln($"Bullet item {i}");
        builder.ListFormat.RemoveNumbers();

        // Iterate through all lists and apply uniform modifications.
        foreach (List list in doc.Lists)
        {
            // Restart numbering at each section for every list.
            list.IsRestartAtEachSection = true;

            // Set the font color of the first level to blue, if the list has levels.
            if (list.ListLevels.Count > 0)
                list.ListLevels[0].Font.Color = Color.Blue;
        }

        // Save the modified document.
        doc.Save("ModifiedLists.docx");
    }
}
