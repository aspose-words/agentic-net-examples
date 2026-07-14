using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a numbered list with three items.
        List numberedList = doc.Lists.Add(ListTemplate.NumberDefault);
        builder.ListFormat.List = numberedList;
        builder.Writeln("Numbered item 1");
        builder.Writeln("Numbered item 2");
        builder.Writeln("Numbered item 3");
        builder.ListFormat.RemoveNumbers();

        // Add a bulleted list with three items.
        List bulletedList = doc.Lists.Add(ListTemplate.BulletDefault);
        builder.ListFormat.List = bulletedList;
        builder.Writeln("Bulleted item A");
        builder.Writeln("Bulleted item B");
        builder.Writeln("Bulleted item C");
        builder.ListFormat.RemoveNumbers();

        // Iterate through all list definitions in the document.
        // For this example we set every list to restart numbering at each section
        // and change the font color of the first level to blue.
        foreach (List list in doc.Lists)
        {
            // Restart numbering at each section.
            list.IsRestartAtEachSection = true;

            // Modify the first level formatting uniformly.
            if (list.ListLevels.Count > 0)
            {
                list.ListLevels[0].Font.Color = Color.Blue;
            }
        }

        // Save the modified document.
        doc.Save("ModifiedLists.docx");
    }
}
