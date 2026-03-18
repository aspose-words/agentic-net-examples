using System;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class MultiLevelListExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create three distinct list styles.
        // Style 1 – simple bullet list.
        Style bulletStyle = doc.Styles.Add(StyleType.List, "BulletListStyle");
        List bulletDef = bulletStyle.List; // definition of the style
        bulletDef.ListLevels[0].NumberStyle = NumberStyle.Bullet;
        bulletDef.ListLevels[0].NumberFormat = "\u2022"; // standard bullet
        bulletDef.ListLevels[0].Font.Name = "Arial";
        bulletDef.ListLevels[0].Font.Color = System.Drawing.Color.DarkGreen;

        // Style 2 – numbered list (decimal).
        Style numberStyle = doc.Styles.Add(StyleType.List, "NumberListStyle");
        List numberDef = numberStyle.List;
        numberDef.ListLevels[0].NumberStyle = NumberStyle.Arabic;
        numberDef.ListLevels[0].NumberFormat = "%1.";
        numberDef.ListLevels[0].Font.Name = "Times New Roman";
        numberDef.ListLevels[0].Font.Color = System.Drawing.Color.DarkBlue;

        // Style 3 – outline list (uppercase Roman).
        Style outlineStyle = doc.Styles.Add(StyleType.List, "OutlineListStyle");
        List outlineDef = outlineStyle.List;
        outlineDef.ListLevels[0].NumberStyle = NumberStyle.UppercaseRoman;
        outlineDef.ListLevels[0].NumberFormat = "%1)";
        outlineDef.ListLevels[0].Font.Name = "Calibri";
        outlineDef.ListLevels[0].Font.Color = System.Drawing.Color.Maroon;

        // Create list objects that reference the styles.
        List bulletList = doc.Lists.Add(bulletStyle);          // reference to bullet style
        List numberList = doc.Lists.Add(numberStyle);          // reference to number style
        List outlineList = doc.Lists.Add(outlineStyle);        // reference to outline style

        DocumentBuilder builder = new DocumentBuilder(doc);

        // Level 1 – uses the bullet list style.
        builder.ListFormat.List = bulletList;
        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("Level 1 – bullet");

        // Level 2 – uses the numbered list style.
        builder.ListFormat.List = numberList;
        builder.ListFormat.ListLevelNumber = 1; // second level (index 1)
        builder.Writeln("Level 2 – number");

        // Level 3 – uses the outline list style.
        builder.ListFormat.List = outlineList;
        builder.ListFormat.ListLevelNumber = 2; // third level (index 2)
        builder.Writeln("Level 3 – outline");

        // End the list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the document.
        doc.Save("MultiLevelList.docx", SaveFormat.Docx);
    }
}
