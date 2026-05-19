using System;
using Aspose.Words;
using Aspose.Words.Lists;

namespace ListLevelValidator
{
    public class Program
    {
        public static void Main()
        {
            // Create a new document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add a list with the maximum allowed 9 levels.
            List list = doc.Lists.Add(ListTemplate.NumberDefault);
            builder.ListFormat.List = list;

            for (int i = 0; i < 9; i++)
            {
                builder.ListFormat.ListLevelNumber = i; // Levels are 0‑based.
                builder.Writeln($"Level {i + 1}");
            }

            builder.ListFormat.RemoveNumbers();

            // Validate that each list in the document does not exceed nine levels.
            bool allValid = true;
            foreach (List lst in doc.Lists)
            {
                int levelCount = lst.ListLevels.Count; // Gets the number of levels in this list.
                if (levelCount > 9)
                {
                    allValid = false;
                    Console.WriteLine($"List ID {lst.ListId} has {levelCount} levels, which exceeds the allowed maximum of 9.");
                }
                else
                {
                    Console.WriteLine($"List ID {lst.ListId} has {levelCount} levels, which is within the allowed limit.");
                }
            }

            Console.WriteLine(allValid
                ? "All lists are valid."
                : "One or more lists exceed the maximum allowed levels.");

            // Save the document.
            doc.Save("ValidatedLists.docx");
        }
    }
}
