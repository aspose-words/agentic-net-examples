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
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a numbered list with three items.
        List numberedList = doc.Lists.Add(ListTemplate.NumberDefault);
        builder.ListFormat.List = numberedList;
        for (int i = 1; i <= 3; i++)
        {
            builder.Writeln($"Numbered item {i}");
        }
        builder.ListFormat.RemoveNumbers();

        // Add a bulleted list with two items.
        List bulletedList = doc.Lists.Add(ListTemplate.BulletDefault);
        builder.ListFormat.List = bulletedList;
        for (int i = 1; i <= 2; i++)
        {
            builder.Writeln($"Bulleted item {i}");
        }
        builder.ListFormat.RemoveNumbers();

        // Iterate through all list definitions in the document and modify them uniformly.
        foreach (List list in doc.Lists)
        {
            // Restart numbering at each section for every list.
            list.IsRestartAtEachSection = true;

            // Change the appearance of the first level of each list.
            if (list.ListLevels.Count > 0)
            {
                list.ListLevels[0].Font.Color = Color.Blue;
                list.ListLevels[0].Font.Size = 14;
            }
        }

        // Save the modified document.
        doc.Save("ModifiedLists.docx");
    }
}
