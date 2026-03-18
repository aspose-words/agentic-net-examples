using System;
using Aspose.Words;
using Aspose.Words.Lists;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Path to the source document.
        const string inputPath = @"C:\Docs\Input.docx";
        // Path where the modified document will be saved.
        const string outputPath = @"C:\Docs\Output.docx";

        // The ListId of the list we want to modify.
        // This value should correspond to an existing list in the document.
        const int targetListId = 2;

        // Load the document from disk.
        Document doc = new Document(inputPath);

        // Retrieve the list with the specified ListId using the GetListByListId method.
        // This method returns null if the list is not found, so we check for that.
        List list = doc.Lists.GetListByListId(targetListId);
        if (list == null)
        {
            Console.WriteLine($"List with ListId {targetListId} was not found.");
            return;
        }

        // Example: modify properties of the first level (index 0) of the list.
        // You can adjust any level (0‑8) in a similar way.
        ListLevel level = list.ListLevels[0];

        // Change the font color of the list label.
        level.Font.Color = Color.Green;

        // Set the alignment of the list number/bullet.
        level.Alignment = ListLevelAlignment.Left;

        // Change the starting number for this level.
        level.StartAt = 5;

        // Optionally, modify other properties such as indentation.
        level.NumberPosition = -18;   // Position of the number/bullet.
        level.TextPosition = 36;      // Position of the text after the number/bullet.
        level.TabPosition = 36;       // Tab stop after the number/bullet.

        // Save the modified document.
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to {outputPath}");
    }
}
