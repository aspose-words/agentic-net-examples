using System;
using System.Linq; // For ElementAt extension method
using Aspose.Words;
using Aspose.Words.Markup;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Insert a drop‑down list StructuredDocumentTag (content control) into the document.
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.DropDownList, MarkupLevel.Block);
        doc.FirstSection.Body.AppendChild(sdt);

        // Populate the list with several items.
        SdtListItemCollection items = sdt.ListItems;
        items.Add(new SdtListItem("Item 1"));
        items.Add(new SdtListItem("Item 2"));
        items.Add(new SdtListItem("Item 3"));
        items.Add(new SdtListItem("Item 4"));

        // Use LINQ's ElementAt to fetch the third item (zero‑based index 2) without iterating.
        SdtListItem thirdItem = items.ElementAt(2); // Equivalent to items[2]

        // Insert the display text of the third item into the document so we can see it.
        // Here we simply add a new paragraph after the content control.
        Paragraph para = new Paragraph(doc);
        para.AppendChild(new Run(doc, $"Third item: {thirdItem.DisplayText}"));
        doc.FirstSection.Body.AppendChild(para);

        // Save the document.
        doc.Save("ThirdItemDemo.docx");
    }
}
