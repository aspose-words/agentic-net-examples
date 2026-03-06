using Aspose.Words;
using Aspose.Words.Lists;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Load an existing DOCX file.
        Document doc = new Document("Input.docx");

        // Obtain a list to modify. If the document has no lists, create a default numbered list.
        List list = doc.Lists.Count > 0 ? doc.Lists[0] : doc.Lists.Add(ListTemplate.NumberDefault);

        // Access the first list level (level 0).
        ListLevel level = list.ListLevels[0];

        // Apply desired formatting to this list level.
        level.Font.Color = Color.Red;                     // Red font color for the label.
        level.Font.Size = 24;                             // Font size 24 points.
        level.Alignment = ListLevelAlignment.Right;       // Align the label to the right of the number position.
        level.NumberStyle = NumberStyle.Ordinal;          // Use ordinal numbering (1., 2., 3., ...).
        level.StartAt = 1;                                // Start numbering at 1.
        level.NumberFormat = "\x0000.";                   // Simple number followed by a period.
        level.NumberPosition = -36;                       // Position of the number/bullet (negative moves it left).
        level.TextPosition = 144;                         // Position where the text starts.
        level.TabPosition = 144;                          // Tab stop after the number.
        level.TrailingCharacter = ListTrailingCharacter.Tab; // Insert a tab after the number.

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
