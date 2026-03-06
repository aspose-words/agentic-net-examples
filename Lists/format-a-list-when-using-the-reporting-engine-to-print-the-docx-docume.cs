using System;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Drawing;
using System.Drawing;

class ListFormattingExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Obtain a DocumentBuilder to insert and format content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ------------------------------------------------------------
        // 1. Create a multi‑level numbered list based on a built‑in template.
        // ------------------------------------------------------------
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // ------------------------------------------------------------
        // 2. Customize the first two levels of the list.
        // ------------------------------------------------------------
        // Level 0 – red, large Arabic numbers.
        ListLevel level0 = list.ListLevels[0];
        level0.Font.Color = Color.Red;
        level0.Font.Size = 14;
        level0.NumberStyle = NumberStyle.Arabic;
        level0.StartAt = 1;               // start numbering at 1
        level0.NumberFormat = "%1.";      // default format, kept for clarity

        // Level 1 – blue, lower‑case letters.
        ListLevel level1 = list.ListLevels[1];
        level1.Font.Color = Color.Blue;
        level1.Font.Size = 12;
        level1.NumberStyle = NumberStyle.LowercaseLetter;
        level1.StartAt = 1;
        level1.NumberFormat = "%2)";      // e.g., a), b), ...

        // ------------------------------------------------------------
        // 3. Write some paragraphs and apply the list formatting.
        // ------------------------------------------------------------
        builder.Writeln("Project Tasks:");

        // Apply the list to subsequent paragraphs.
        builder.ListFormat.List = list;

        // First level items.
        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("Design phase");
        builder.Writeln("Implementation phase");

        // Indent to second level for sub‑tasks.
        builder.ListFormat.ListLevelNumber = 1;
        builder.Writeln("Write code");
        builder.Writeln("Unit tests");

        // Return to first level.
        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("Testing phase");
        builder.Writeln("Deployment phase");

        // ------------------------------------------------------------
        // 4. End the list formatting.
        // ------------------------------------------------------------
        builder.ListFormat.RemoveNumbers();

        // ------------------------------------------------------------
        // 5. Save the document (programmatic printing can be done later).
        // ------------------------------------------------------------
        doc.Save("FormattedList.docx");
    }
}
