using System;
using Aspose.Words;
using Aspose.Words.Lists;

namespace ListLevelValidator
{
    public class Program
    {
        public static void Main()
        {
            // Create a new empty document.
            Document doc = new Document();

            // Add a sample list using a predefined template.
            // All lists created by Aspose.Words contain up to 9 levels.
            List sampleList = doc.Lists.Add(ListTemplate.NumberDefault);

            // Validate each list in the document.
            foreach (List list in doc.Lists)
            {
                int levelCount = list.ListLevels.Count;

                // According to the API, a list may have 1 to 9 levels.
                if (levelCount > 9)
                {
                    Console.WriteLine($"List ID {list.ListId} exceeds the maximum allowed levels: {levelCount}");
                }
                else
                {
                    Console.WriteLine($"List ID {list.ListId} is valid with {levelCount} level(s).");
                }
            }

            // Save the document (optional, just to demonstrate the save lifecycle).
            doc.Save("ValidatedLists.docx");
        }
    }
}
