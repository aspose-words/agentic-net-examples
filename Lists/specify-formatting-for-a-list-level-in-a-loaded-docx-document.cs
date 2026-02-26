using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Lists;

class Program
{
    static void Main()
    {
        // Load an existing DOCX document.
        Document doc = new Document("input.docx");

        // Ensure the document contains at least one list.
        if (doc.Lists.Count == 0)
        {
            Console.WriteLine("The document does not contain any lists.");
            return;
        }

        // Get the first list in the document.
        List list = doc.Lists[0];

        // Access the first level (level 0) of the list.
        ListLevel level = list.ListLevels[0];

        // Example formatting changes for this list level:
        // Set the number style to uppercase Roman numerals.
        level.NumberStyle = NumberStyle.UppercaseRoman;

        // Define a custom number format with a prefix and suffix.
        // "\x0000" is a placeholder for the current level number.
        level.NumberFormat = "Section \x0000:";

        // Align the number to the right of the number position.
        level.Alignment = ListLevelAlignment.Right;

        // Set the font used for the list label.
        level.Font.Name = "Arial";
        level.Font.Size = 12;
        level.Font.Color = Color.DarkBlue;
        level.Font.Bold = true;

        // Adjust indent positions (in points).
        level.NumberPosition = -18;   // Position of the number/bullet.
        level.TextPosition = 36;      // Position of the text after the number.
        level.TabPosition = 36;       // Tab stop after the number.

        // Save the modified document.
        doc.Save("output.docx");
    }
}
