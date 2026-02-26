using System;
using Aspose.Words;
using Aspose.Words.Markup;

class InsertContentControls
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a DocumentBuilder attached to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a plain‑text content control (StructuredDocumentTag) at the current cursor position.
        // The method returns the newly created StructuredDocumentTag node.
        StructuredDocumentTag plainTextControl = builder.InsertStructuredDocumentTag(SdtType.PlainText);

        // Set optional properties such as Title and Tag for identification.
        plainTextControl.Title = "CustomerName";
        plainTextControl.Tag = "CustomerNameTag";

        // Write some placeholder text inside the content control.
        builder.Write("Enter customer name here");

        // Move the cursor out of the content control to continue normal document editing.
        builder.Writeln(); // Ends the paragraph inside the control.
        builder.MoveToDocumentEnd(); // Ensure the cursor is after the control.

        // Insert a second content control – a repeating section (useful for tables or lists).
        StructuredDocumentTag repeatingSection = builder.InsertStructuredDocumentTag(SdtType.RepeatingSection);
        repeatingSection.Title = "OrderItems";
        repeatingSection.Tag = "OrderItemsTag";

        // Inside the repeating section, insert a table with a single row as a template.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Item");
        builder.InsertCell();
        builder.Write("Quantity");
        builder.EndRow();
        builder.EndTable();

        // Finish the document.
        builder.Writeln();

        // Save the document to a file (WordprocessingML format – .docx).
        doc.Save("ContentControls.docx");
    }
}
