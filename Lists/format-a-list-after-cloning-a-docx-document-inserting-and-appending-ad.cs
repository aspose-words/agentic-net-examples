using System;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the original DOCX document.
        Document srcDoc = new Document("Original.docx");

        // Clone the original document to create an independent copy.
        Document clonedDoc = srcDoc.Clone();

        // Create a new numbered list in the cloned document and apply custom formatting.
        List list = clonedDoc.Lists.Add(ListTemplate.NumberArabicParenthesis);
        list.ListLevels[0].Font.Color = System.Drawing.Color.Red;
        list.ListLevels[0].Font.Size = 12;
        list.ListLevels[0].Alignment = ListLevelAlignment.Right;

        // Add list items to the cloned document using DocumentBuilder.
        DocumentBuilder builder = new DocumentBuilder(clonedDoc);
        builder.Writeln("List starts:");
        builder.ListFormat.List = list;
        for (int i = 1; i <= 5; i++)
        {
            builder.Writeln($"Item {i}");
        }
        builder.ListFormat.RemoveNumbers(); // End list formatting.

        // Insert additional content at the end of the original document.
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);
        srcBuilder.MoveToDocumentEnd();
        srcBuilder.InsertBreak(BreakType.PageBreak);
        srcBuilder.Writeln("Additional content inserted before appending cloned document.");

        // Append the cloned document to the original, preserving source list numbering.
        ImportFormatOptions importOptions = new ImportFormatOptions
        {
            KeepSourceNumbering = true // Keep the numbering from the cloned document.
        };
        srcDoc.AppendDocument(clonedDoc, ImportFormatMode.KeepSourceFormatting, importOptions);
        srcDoc.UpdateListLabels(); // Refresh list numbers after appending.

        // Split the combined document into two parts:
        //   - First part: pages 1 to 2.
        //   - Second part: pages 3 to the end.
        Document firstPart = srcDoc.ExtractPages(1, 2);
        Document secondPart = srcDoc.ExtractPages(3, srcDoc.PageCount);

        // Save the resulting documents.
        srcDoc.Save("Combined.docx");
        firstPart.Save("Combined_Part1.docx");
        secondPart.Save("Combined_Part2.docx");
    }
}
