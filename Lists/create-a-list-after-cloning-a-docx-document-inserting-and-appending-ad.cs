using System;
using Aspose.Words;
using Aspose.Words.Lists;

class DocumentProcessing
{
    static void Main()
    {
        // Load the original DOCX document.
        Document srcDoc = new Document("input.docx");

        // Clone the original document (deep copy).
        Document clonedDoc = srcDoc.Clone();

        // Insert additional content into the cloned document.
        DocumentBuilder builder = new DocumentBuilder(clonedDoc);
        builder.Writeln("Inserted paragraph in cloned document.");

        // Create a numbered list in the cloned document.
        List list = clonedDoc.Lists.Add(ListTemplate.NumberDefault);
        builder.ListFormat.List = list;
        for (int i = 1; i <= 5; i++)
        {
            builder.Writeln($"List item {i}");
        }
        // End list formatting.
        builder.ListFormat.RemoveNumbers();

        // Append the modified cloned document to the end of the original document.
        srcDoc.AppendDocument(clonedDoc, ImportFormatMode.KeepSourceFormatting);

        // Split the combined document into two parts:
        // 1) First page.
        Document firstPage = srcDoc.ExtractPages(1, 1);
        // 2) Remaining pages (from page 2 to the end).
        Document rest = srcDoc.ExtractPages(2, srcDoc.PageCount);

        // Save the resulting documents.
        srcDoc.Save("Combined.docx");
        firstPage.Save("FirstPage.docx");
        rest.Save("Rest.docx");
    }
}
