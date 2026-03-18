using System;
using Aspose.Words;
using Aspose.Words.Lists;

class ListLevelValidator
{
    static void Main()
    {
        // Load an existing Word document.
        Document doc = new Document("Input.docx");

        // Flag to indicate whether all lists are within the allowed level range.
        bool allListsValid = true;

        // Iterate through every list in the document.
        foreach (List list in doc.Lists)
        {
            // The ListLevelCollection.Count property gives the number of levels defined for this list.
            int levelCount = list.ListLevels.Count;

            // According to Aspose.Words, a list may contain up to 9 levels (0‑8).
            if (levelCount > 9)
            {
                allListsValid = false;
                Console.WriteLine($"List ID {list.ListId} has {levelCount} levels, which exceeds the maximum of 9.");
            }
            else
            {
                Console.WriteLine($"List ID {list.ListId} is valid with {levelCount} level(s).");
            }
        }

        // Report overall validation result.
        if (allListsValid)
        {
            Console.WriteLine("All lists in the document are within the allowed 9‑level limit.");
        }
        else
        {
            Console.WriteLine("One or more lists exceed the 9‑level limit.");
        }

        // Save the (unchanged) document – this uses the standard save lifecycle.
        doc.Save("ValidatedOutput.docx");
    }
}
