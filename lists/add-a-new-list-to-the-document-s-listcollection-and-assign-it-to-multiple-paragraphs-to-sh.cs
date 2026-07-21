using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add some paragraphs.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("First paragraph (no list).");
        builder.Writeln("Second paragraph (will share list).");
        builder.Writeln("Third paragraph (will share list).");
        builder.Writeln("Fourth paragraph (no list).");

        // Create a new list based on a predefined template (bulleted list).
        List sharedList = doc.Lists.Add(ListTemplate.BulletDefault);

        // Retrieve all paragraphs in the document.
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

        // Assign the shared list to the second and third paragraphs.
        // Set the list level number to 0 (first level) for both.
        int index = 0;
        foreach (Paragraph para in paragraphs.OfType<Paragraph>())
        {
            if (index == 1 || index == 2) // second and third paragraphs (0‑based)
            {
                para.ListFormat.List = sharedList;
                para.ListFormat.ListLevelNumber = 0;
            }
            index++;
        }

        // Save the document to a file in the current directory.
        doc.Save("SharedListExample.docx");
    }
}
