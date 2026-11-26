using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define first list style (numeric with sublevel as lower letters).
        List list1 = doc.Lists.Add(ListTemplate.NumberDefault);
        ListLevel level0 = list1.ListLevels[0];
        level0.NumberStyle = NumberStyle.Arabic;
        level0.NumberFormat = "%1.";
        level0.Alignment = ListLevelAlignment.Left;
        level0.Font.Name = "Arial";
        level0.Font.Size = 12;

        ListLevel level1 = list1.ListLevels[1];
        level1.NumberStyle = NumberStyle.LowercaseLetter;   // corrected enum name
        level1.NumberFormat = "%2)";
        level1.Alignment = ListLevelAlignment.Left;
        level1.Font.Name = "Arial";
        level1.Font.Size = 12;

        // Define second list style (bulleted with sublevel as upper Roman numerals).
        List list2 = doc.Lists.Add(ListTemplate.BulletDefault);
        ListLevel bLevel0 = list2.ListLevels[0];
        bLevel0.NumberStyle = NumberStyle.Bullet;
        bLevel0.NumberFormat = "\u2022"; // bullet character
        bLevel0.Font.Name = "Calibri";
        bLevel0.Font.Size = 12;

        ListLevel bLevel1 = list2.ListLevels[1];
        bLevel1.NumberStyle = NumberStyle.UppercaseRoman;   // corrected enum name
        bLevel1.NumberFormat = "%2.";
        bLevel1.Font.Name = "Calibri";
        bLevel1.Font.Size = 12;

        // Insert items using the first list style.
        builder.ListFormat.List = list1;
        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("Item 1 - List1 Level 0");
        builder.ListFormat.ListLevelNumber = 1;
        builder.Writeln("Item 1.1 - List1 Level 1");
        builder.ListFormat.RemoveNumbers();

        // Insert items using the second list style.
        builder.ListFormat.List = list2;
        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("Item A - List2 Level 0");
        builder.ListFormat.ListLevelNumber = 1;
        builder.Writeln("Item A.I - List2 Level 1");
        builder.ListFormat.RemoveNumbers();

        // Save the document.
        string fileName = "MultiLevelList.docx";
        doc.Save(fileName);

        // Verify the file was created.
        Console.WriteLine($"Document saved to: {Path.GetFullPath(fileName)}");
        Console.WriteLine(File.Exists(fileName) ? "File exists." : "File was not created.");
    }
}
