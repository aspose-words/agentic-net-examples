using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Add a new list to the document using a built‑in template.
        // The Add method returns the created List object.
        List list = doc.Lists.Add(ListTemplate.BulletDefault);

        // Attempt to access more than the allowed nine list levels.
        // The ListLevels collection supports indices 0‑8; index 9 throws ArgumentOutOfRangeException.
        for (int i = 0; i < 10; i++)
        {
            try
            {
                // This will throw ArgumentOutOfRangeException when i >= 9.
                ListLevel level = list.ListLevels[i];

                // Optional configuration of the level – shown for demonstration.
                level.NumberStyle = NumberStyle.Arabic;
                level.NumberPosition = 0;
                level.Alignment = ListLevelAlignment.Left;
                level.Font.Name = "Arial";
                level.Font.Size = 12;
            }
            catch (ArgumentOutOfRangeException ex)
            {
                // Handle the exception for exceeding the maximum number of levels.
                Console.WriteLine($"Exception caught for level {i}: {ex.Message}");
            }
        }

        // Add a paragraph that uses the list to keep the document valid.
        Paragraph para = new Paragraph(doc);
        para.ListFormat.List = list;
        para.ListFormat.ListLevelNumber = 0; // First level (0‑based index).
        para.AppendChild(new Run(doc, "Sample list item"));
        doc.FirstSection.Body.AppendChild(para);

        // Save the document.
        doc.Save("Result.docx");
    }
}
