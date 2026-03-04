using System;
using Aspose.Words;
using Aspose.Words.Markup;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a plain‑text content control.
        StructuredDocumentTag plain = builder.InsertStructuredDocumentTag(SdtType.PlainText);
        plain.Title = "PlainTextControl";
        plain.Tag = "PlainTag";
        builder.Writeln("Plain text inside control.");

        // Insert a rich‑text content control.
        StructuredDocumentTag rich = builder.InsertStructuredDocumentTag(SdtType.RichText);
        rich.Title = "RichTextControl";
        rich.Tag = "RichTag";
        builder.Font.Bold = true;
        builder.Font.Size = 14;
        builder.Writeln("Bold rich text.");
        // Reset formatting for subsequent text.
        builder.Font.Bold = false;
        builder.Font.Size = 12;
        builder.Writeln();

        // Insert a checkbox content control.
        StructuredDocumentTag checkBox = builder.InsertStructuredDocumentTag(SdtType.Checkbox);
        checkBox.Title = "AcceptTerms";
        checkBox.Tag = "TermsCheck";
        checkBox.Checked = true; // default state
        builder.Writeln("I accept the terms and conditions.");
        builder.Writeln();

        // Insert a drop‑down list content control.
        StructuredDocumentTag dropDown = builder.InsertStructuredDocumentTag(SdtType.DropDownList);
        dropDown.Title = "CountrySelector";
        dropDown.Tag = "CountryTag";
        dropDown.ListItems.Add(new SdtListItem("United States", "US"));
        dropDown.ListItems.Add(new SdtListItem("Canada", "CA"));
        dropDown.ListItems.Add(new SdtListItem("United Kingdom", "UK"));
        builder.Writeln("Select your country: ");

        // Insert a combo‑box content control.
        StructuredDocumentTag comboBox = builder.InsertStructuredDocumentTag(SdtType.ComboBox);
        comboBox.Title = "ColorSelector";
        comboBox.Tag = "ColorTag";
        comboBox.ListItems.Add(new SdtListItem("Red", "R"));
        comboBox.ListItems.Add(new SdtListItem("Green", "G"));
        comboBox.ListItems.Add(new SdtListItem("Blue", "B"));
        builder.Writeln("Choose a color: ");
        builder.Writeln();

        // Insert a picture content control.
        StructuredDocumentTag picture = builder.InsertStructuredDocumentTag(SdtType.Picture);
        picture.Title = "Signature";
        picture.Tag = "SignatureTag";
        // Insert an image inside the picture control (ensure the file exists).
        builder.InsertImage("sample.png");
        builder.Writeln();

        // Insert a date content control.
        StructuredDocumentTag date = builder.InsertStructuredDocumentTag(SdtType.Date);
        date.Title = "DateOfBirth";
        date.Tag = "DOBTag";
        date.DateDisplayFormat = "yyyy-MM-dd";
        builder.Writeln("Date of Birth: ");

        // Save the document as DOCX.
        doc.Save("ContentControls.docx");
    }
}
