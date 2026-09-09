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

        // Add a numbered list and some items.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);
        builder.ListFormat.List = list;
        for (int i = 1; i <= 3; i++)
        {
            builder.Writeln($"Item {i}");
        }
        builder.ListFormat.RemoveNumbers();

        // Iterate through all list definitions in the document.
        foreach (List lst in doc.Lists)
        {
            // Example modification: restart numbering at each section.
            lst.IsRestartAtEachSection = true;

            // Example modification: set the font of the first level to green and bold.
            if (lst.ListLevels.Count > 0)
            {
                lst.ListLevels[0].Font.Color = Color.Green;
                lst.ListLevels[0].Font.Bold = true;
            }
        }

        // Save the document to the output file.
        doc.Save("ModifiedLists.docx");
    }
}
