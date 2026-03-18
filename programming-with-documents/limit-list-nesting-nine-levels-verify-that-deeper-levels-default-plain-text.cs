using System;
using System.Diagnostics;
using Aspose.Words;
using Aspose.Words.Lists;

class ListNestingDemo
{
    static void Main()
    {
        // Folder for output files.
        string artifactsDir = "Artifacts";
        System.IO.Directory.CreateDirectory(artifactsDir);

        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a default numbered list (has 9 levels: 0‑8).
        builder.ListFormat.List = doc.Lists.Add(ListTemplate.NumberDefault);

        // Add paragraphs with increasing list levels, including levels beyond the supported range.
        for (int i = 0; i <= 10; i++) // 0‑8 are valid, 9‑10 exceed the limit.
        {
            builder.ListFormat.ListLevelNumber = i;
            builder.Writeln($"Level {i}");
        }

        // End the list.
        builder.ListFormat.RemoveNumbers();

        // Save the document.
        string docPath = System.IO.Path.Combine(artifactsDir, "ListNesting.docx");
        doc.Save(docPath);

        // Verify that levels 0‑8 are formatted as list items,
        // while deeper levels default to plain text (IsListItem == false).
        foreach (Paragraph para in doc.FirstSection.Body.Paragraphs)
        {
            int level = para.ListFormat.ListLevelNumber;
            bool isListItem = para.ListFormat.IsListItem;

            if (level > 8)
            {
                Debug.Assert(!isListItem,
                    $"Paragraph at level {level} should NOT be a list item, but it is.");
            }
            else
            {
                Debug.Assert(isListItem,
                    $"Paragraph at level {level} should be a list item, but it is not.");
            }
        }

        // Optional: output verification result to console.
        Console.WriteLine("List nesting verification completed successfully.");
    }
}
