// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Lists;

class ListFormattingDemo
{
    static void Main()
    {
        // Load the source DOCX document.
        Document doc = new Document("Template.docx");

        // -------------------------------------------------
        // 1. Create a custom bullet list and add sample items.
        // -------------------------------------------------
        List bulletList = doc.Lists.Add(ListTemplate.BulletDefault);
        // Customize the first level of the list (blue bullet, larger font).
        bulletList.ListLevels[0].Font.Color = Color.Blue;
        bulletList.ListLevels[0].Font.Size = 12;

        DocumentBuilder builder = new DocumentBuilder(doc);
        // Apply the list to subsequent paragraphs.
        builder.ListFormat.List = bulletList;
        builder.Writeln("First bullet item");
        builder.Writeln("Second bullet item");
        // End the list.
        builder.ListFormat.RemoveNumbers();

        // -------------------------------------------------
        // 2. Perform a simple mail merge that will use the list.
        // -------------------------------------------------
        // Assume the template contains a MERGEFIELD called "Item".
        // The mail merge data will be a single row with two items.
        string[] fieldNames = { "Item" };
        object[] fieldValues = { "Merged list item" };
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // After mail merge, apply the same list formatting to the merged paragraph.
        Paragraph mergedParagraph = doc.GetChildNodes(NodeType.Paragraph, true)
                                      .Cast<Paragraph>()
                                      .FirstOrDefault(p => p.GetText().Contains("Merged list item"));
        if (mergedParagraph != null)
        {
            mergedParagraph.ListFormat.List = bulletList;
            mergedParagraph.ListFormat.ListLevelNumber = 0;
        }

        // -------------------------------------------------
        // 3. LINQ reporting: find paragraphs containing a keyword and format them as a numbered list.
        // -------------------------------------------------
        List numberedList = doc.Lists.Add(ListTemplate.NumberDefault);
        // Customize the first level (red numbers).
        numberedList.ListLevels[0].Font.Color = Color.Red;
        numberedList.ListLevels[0].Font.Size = 12;

        var reportParagraphs = doc.GetChildNodes(NodeType.Paragraph, true)
                                  .Cast<Paragraph>()
                                  .Where(p => p.GetText().Contains("[Report]"))
                                  .ToList();

        foreach (var para in reportParagraphs)
        {
            para.ListFormat.List = numberedList;
            para.ListFormat.ListLevelNumber = 0;
        }

        // -------------------------------------------------
        // 4. Print the document (to the default printer).
        // -------------------------------------------------
        // The Print method uses the system's default printer.
        doc.Print();

        // -------------------------------------------------
        // 5. Save the modified document to PDF (rendering).
        // -------------------------------------------------
        doc.Save("FormattedOutput.pdf");
    }
}
