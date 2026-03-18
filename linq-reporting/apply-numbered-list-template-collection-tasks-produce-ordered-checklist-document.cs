using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Lists;

public class ChecklistGenerator
{
    // Generates a Word document with an ordered checklist from the supplied tasks.
    public static void CreateChecklist(IEnumerable<string> tasks, string outputPath)
    {
        // 1. Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // 2. Create a numbered list based on the default numbering template (1., 2., 3., ...).
        List numberedList = doc.Lists.Add(ListTemplate.NumberDefault);

        // 3. Apply the list to each task.
        foreach (string task in tasks)
        {
            // Assign the list to the current paragraph.
            builder.ListFormat.List = numberedList;
            // Ensure we are using the first level of the list (0‑based index).
            builder.ListFormat.ListLevelNumber = 0;
            // Write the task text as a list item.
            builder.Writeln(task);
        }

        // 4. End list formatting for any subsequent content.
        builder.ListFormat.RemoveNumbers();

        // 5. Save the document to the specified file.
        doc.Save(outputPath);
    }

    // Example usage.
    public static void Main()
    {
        var tasks = new List<string>
        {
            "Gather requirements",
            "Design architecture",
            "Implement features",
            "Write unit tests",
            "Perform code review",
            "Deploy to production"
        };

        string outputFile = "Checklist.docx";
        CreateChecklist(tasks, outputFile);
        Console.WriteLine($"Checklist saved to '{outputFile}'.");
    }
}
