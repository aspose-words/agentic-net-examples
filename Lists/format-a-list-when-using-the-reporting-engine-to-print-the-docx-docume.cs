using Aspose.Words;
using Aspose.Words.Lists;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Create a new document and a builder to insert content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -----------------------------------------------------------------
        // 1. Create a list based on a predefined template.
        // -----------------------------------------------------------------
        // Use the ListCollection.Add method (provided by Aspose.Words) to
        // create a numbered list that uses uppercase letters (A., B., C., ...).
        List list = doc.Lists.Add(ListTemplate.NumberUppercaseLetterDot);

        // -----------------------------------------------------------------
        // 2. Customize the first level of the list.
        // -----------------------------------------------------------------
        ListLevel level0 = list.ListLevels[0];
        level0.Font.Color = Color.DarkBlue;          // Bullet/number color.
        level0.Font.Size = 14;                       // Font size.
        level0.NumberStyle = NumberStyle.UppercaseLetter;
        level0.StartAt = 1;                          // Start numbering at A.
        level0.NumberFormat = "%1.";                 // Format: A., B., ...
        level0.NumberPosition = -18;                // Position of the number.
        level0.TextPosition = 36;                    // Position of the text.
        level0.TabPosition = 36;                     // Tab stop after the number.

        // -----------------------------------------------------------------
        // 3. Apply the list to paragraphs.
        // -----------------------------------------------------------------
        builder.Writeln("Key Advantages:");
        builder.ListFormat.List = list;              // Attach the list to the builder.
        builder.ListFormat.ListLevelNumber = 0;      // Use first level.

        builder.Writeln("High performance");
        builder.Writeln("Cross‑platform support");
        builder.Writeln("Rich API");

        // -----------------------------------------------------------------
        // 4. Create a sub‑list (second level) to demonstrate nesting.
        // -----------------------------------------------------------------
        builder.ListFormat.ListIndent();             // Increase level to 1 (second level).
        builder.Writeln("Detailed point A");
        builder.Writeln("Detailed point B");
        builder.ListFormat.ListOutdent();            // Return to first level.

        // -----------------------------------------------------------------
        // 5. End list formatting.
        // -----------------------------------------------------------------
        builder.ListFormat.RemoveNumbers();          // Remove list formatting from subsequent paragraphs.

        // -----------------------------------------------------------------
        // 6. Save the document (can be printed later via reporting engine or UI).
        // -----------------------------------------------------------------
        doc.Save("FormattedList.docx");
    }
}
