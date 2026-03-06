using Aspose.Words;
using Aspose.Words.Markup;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder attached to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a plain‑text content control (SDT).
        StructuredDocumentTag plainText = builder.InsertStructuredDocumentTag(SdtType.PlainText);
        plainText.Title = "PlainTextControl";
        plainText.Tag = "PlainTextTag";
        builder.Writeln(); // Move cursor after the control.

        // Insert a rich‑text content control.
        StructuredDocumentTag richText = builder.InsertStructuredDocumentTag(SdtType.RichText);
        richText.Title = "RichTextControl";
        richText.Tag = "RichTextTag";
        builder.Writeln();

        // Insert a checkbox content control.
        StructuredDocumentTag checkBox = builder.InsertStructuredDocumentTag(SdtType.Checkbox);
        checkBox.Title = "AcceptTerms";
        checkBox.Tag = "AcceptTermsTag";
        checkBox.Checked = false; // Default unchecked.
        builder.Writeln();

        // Insert a dropdown list content control and populate its items.
        StructuredDocumentTag dropDown = builder.InsertStructuredDocumentTag(SdtType.DropDownList);
        dropDown.Title = "CountrySelection";
        dropDown.Tag = "CountryTag";
        dropDown.ListItems.Add(new SdtListItem("USA", "USA"));
        dropDown.ListItems.Add(new SdtListItem("Canada", "Canada"));
        dropDown.ListItems.Add(new SdtListItem("Mexico", "Mexico"));
        builder.Writeln();

        // Save the document with the inserted content controls.
        doc.Save("ContentControls.docx");
    }
}
