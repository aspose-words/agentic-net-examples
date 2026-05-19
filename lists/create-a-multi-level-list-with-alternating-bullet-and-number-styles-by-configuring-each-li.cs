using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a multilevel list based on a default template.
        List multiLevelList = doc.Lists.Add(ListTemplate.NumberDefault);

        // Configure each level: even levels -> bullet, odd levels -> Arabic numbers.
        for (int i = 0; i < multiLevelList.ListLevels.Count; i++)
        {
            ListLevel level = multiLevelList.ListLevels[i];
            level.Font.Name = "Arial";
            level.Font.Size = 12;
            level.TrailingCharacter = ListTrailingCharacter.Tab;

            if (i % 2 == 0) // Bullet level
            {
                level.NumberStyle = NumberStyle.Bullet;
                // Unicode bullet character.
                level.NumberFormat = "\u2022";
            }
            else // Numbered level
            {
                level.NumberStyle = NumberStyle.Arabic;
                // Show the current level number followed by a dot.
                level.NumberFormat = "\x0000.";
            }
        }

        // Apply the list to the builder.
        builder.ListFormat.List = multiLevelList;

        // Build a sample list with several levels.
        builder.Writeln("Level 0 – Bullet");
        builder.ListFormat.ListIndent(); // Level 1 – Number
        builder.Writeln("Level 1 – Number");
        builder.ListFormat.ListIndent(); // Level 2 – Bullet
        builder.Writeln("Level 2 – Bullet");
        builder.ListFormat.ListIndent(); // Level 3 – Number
        builder.Writeln("Level 3 – Number");
        builder.ListFormat.ListOutdent(); // Back to Level 2
        builder.Writeln("Another Level 2 – Bullet");
        builder.ListFormat.ListOutdent(); // Back to Level 1
        builder.Writeln("Another Level 1 – Number");
        builder.ListFormat.ListOutdent(); // Back to Level 0
        builder.Writeln("Another Level 0 – Bullet");

        // End the list.
        builder.ListFormat.RemoveNumbers();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "MultiLevelList.docx");
        doc.Save(outputPath);
    }
}
