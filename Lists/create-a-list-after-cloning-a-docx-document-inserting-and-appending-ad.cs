using System;
using Aspose.Words;
using Aspose.Words.Lists;

class Program
{
    static void Main()
    {
        // Load the original document.
        Document srcDoc = new Document("Source.docx");

        // Clone the original document (deep copy).
        Document clonedDoc = srcDoc.Clone();

        // Create a numbered list in the cloned document.
        DocumentBuilder listBuilder = new DocumentBuilder(clonedDoc);
        List numberedList = clonedDoc.Lists.Add(ListTemplate.NumberDefault);
        listBuilder.ListFormat.List = numberedList;
        for (int i = 1; i <= 5; i++)
        {
            listBuilder.Writeln($"Item {i}");
        }
        // End list formatting.
        listBuilder.ListFormat.RemoveNumbers();

        // Create an additional document with extra content.
        Document extraDoc = new Document();
        DocumentBuilder extraBuilder = new DocumentBuilder(extraDoc);
        extraBuilder.Writeln("Additional content appended to the cloned document.");

        // Append the extra document to the cloned document.
        clonedDoc.AppendDocument(extraDoc, ImportFormatMode.KeepSourceFormatting);

        // Insert the original source document at the end of the cloned document.
        DocumentBuilder insertBuilder = new DocumentBuilder(clonedDoc);
        insertBuilder.MoveToDocumentEnd();
        insertBuilder.InsertBreak(BreakType.PageBreak);
        insertBuilder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

        // Split the cloned document into two parts: pages 1‑2 and the remaining pages.
        Document part1 = clonedDoc.ExtractPages(1, 2);
        Document part2 = clonedDoc.ExtractPages(3, clonedDoc.PageCount);

        // Save the resulting documents.
        clonedDoc.Save("Result.docx");
        part1.Save("Result_Part1.docx");
        part2.Save("Result_Part2.docx");
    }
}
