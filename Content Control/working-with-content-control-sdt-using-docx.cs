using System;
using Aspose.Words;
using Aspose.Words.Markup;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a plain‑text content control (SDT) inline.
        StructuredDocumentTag nameSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "CustomerName",          // Friendly name.
            Tag = "CustomerNameTag",         // Identifier used for lookup.
            LockContents = true             // Prevent the user from editing the contents.
        };
        // Write placeholder text that will appear inside the SDT.
        builder.Write("Enter name here");
        // Insert the SDT after the placeholder text.
        builder.InsertNode(nameSdt);

        // Insert a drop‑down list content control.
        StructuredDocumentTag countrySdt = new StructuredDocumentTag(doc, SdtType.DropDownList, MarkupLevel.Inline)
        {
            Title = "Country",
            Tag = "CountryTag"
        };
        // Populate the list with items.
        countrySdt.ListItems.Add(new SdtListItem("USA", "USA"));
        countrySdt.ListItems.Add(new SdtListItem("Canada", "Canada"));
        countrySdt.ListItems.Add(new SdtListItem("Mexico", "Mexico"));
        builder.Write(" Select country: ");
        builder.InsertNode(countrySdt);

        // Save the document containing the content controls.
        doc.Save("ContentControls.docx");

        // Load the document back to demonstrate accessing the SDTs.
        Document loadedDoc = new Document("ContentControls.docx");

        // Enumerate all structured document tags in the document.
        foreach (IStructuredDocumentTag sdt in loadedDoc.Range.StructuredDocumentTags)
        {
            Console.WriteLine($"Title: {sdt.Title}, Tag: {sdt.Tag}, Type: {sdt.SdtType}");
        }

        // Replace the placeholder text inside the first SDT.
        loadedDoc.Range.Replace("Enter name here", "John Doe");

        // Save the updated document.
        loadedDoc.Save("ContentControls_Updated.docx");
    }
}
