using System;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Tables;

namespace ListLevelValidationExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a bulleted list template.
            List list = doc.Lists.Add(ListTemplate.BulletDefault);

            // Build a list with varying nesting levels (including an invalid level > 9).
            for (int level = 1; level <= 10; level++)
            {
                // Assign the list to the builder.
                builder.ListFormat.List = list;
                // Set the desired nesting level.
                builder.ListFormat.ListLevelNumber = level;
                // Write the list item.
                builder.Writeln($"List item at level {level}");
            }

            // Finish the list.
            builder.ListFormat.RemoveNumbers();

            // Validate that no list exceeds nine levels.
            bool hasInvalidLevel = false;
            foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
            {
                if (para.ListFormat.IsListItem)
                {
                    int currentLevel = para.ListFormat.ListLevelNumber;
                    if (currentLevel > 9)
                    {
                        hasInvalidLevel = true;
                        Console.WriteLine($"Invalid list level detected: Paragraph \"{para.GetText().Trim()}\" is at level {currentLevel}.");
                    }
                }
            }

            if (!hasInvalidLevel)
            {
                Console.WriteLine("All list levels are within the allowed range (1-9).");
            }

            // Save the document (optional, demonstrates lifecycle usage).
            doc.Save("ValidatedListDocument.docx");
        }
    }
}
