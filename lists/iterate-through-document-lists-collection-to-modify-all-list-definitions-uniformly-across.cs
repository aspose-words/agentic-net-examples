using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a numbered list and add two items.
        List numberedList = doc.Lists.Add(ListTemplate.NumberDefault);
        builder.ListFormat.List = numberedList;
        builder.Writeln("Numbered Item 1");
        builder.Writeln("Numbered Item 2");
        builder.ListFormat.RemoveNumbers();

        // Create a bullet list and add two items.
        List bulletList = doc.Lists.Add(ListTemplate.BulletDefault);
        builder.ListFormat.List = bulletList;
        builder.Writeln("Bullet Item 1");
        builder.Writeln("Bullet Item 2");
        builder.ListFormat.RemoveNumbers();

        // Iterate through all list definitions in the document and modify them uniformly.
        foreach (List list in doc.Lists)
        {
            // Apply the same formatting to every level of each list.
            foreach (ListLevel level in list.ListLevels)
            {
                // Change the numbering style to uppercase Roman numerals.
                level.NumberStyle = NumberStyle.UppercaseRoman;

                // Optionally, set a different font for the list numbers.
                level.Font.Name = "Arial";
                level.Font.Size = 10;
            }
        }

        // Save the modified document.
        doc.Save("ModifiedLists.docx");

        // Indicate completion.
        Console.WriteLine("Document created and list definitions modified.");
    }
}
