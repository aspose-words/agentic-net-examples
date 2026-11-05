using System;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Tables;
using Aspose.Words.Saving; // Needed for OoxmlSaveOptions and OoxmlCompliance enum

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a list based on a built‑in template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Apply the list to the builder.
        builder.ListFormat.List = list;

        // First top‑level item.
        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("First item");

        // Nested second level item.
        builder.ListFormat.ListLevelNumber = 1;
        builder.Writeln("Nested item");

        // Return to top level for another item.
        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("Second item");

        // End list formatting.
        builder.ListFormat.RemoveNumbers();

        // Validation: count paragraphs that are part of the list.
        int listParagraphCount = 0;
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            if (para.ListFormat.IsListItem)
                listParagraphCount++;
        }

        // Ensure OoxmlCompliance is set higher than Ecma376 to retain custom list settings.
        // Use OoxmlSaveOptions to specify the compliance level.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
        {
            Compliance = OoxmlCompliance.Iso29500_2008_Transitional
        };

        // Save the document with the specified compliance level.
        doc.Save("Output.docx", saveOptions);
    }
}
