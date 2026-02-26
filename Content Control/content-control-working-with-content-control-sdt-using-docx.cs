using System;
using Aspose.Words;
using Aspose.Words.Markup;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a line before the content control.
        builder.Writeln("Below is a checkbox content control:");

        // Create an inline checkbox StructuredDocumentTag (content control).
        StructuredDocumentTag checkBox = new StructuredDocumentTag(doc, SdtType.Checkbox, MarkupLevel.Inline)
        {
            Title = "AcceptTerms",          // Friendly name shown in the UI.
            Tag = "AcceptTermsTag",         // Tag used for identification in code.
            Checked = false,                // Initial state of the checkbox.
            IsShowingPlaceholderText = true // Show placeholder when the control is empty.
        };

        // Define custom symbols for the checked and unchecked states.
        // 0x2611 = ☑, 0x2610 = ☐ (Unicode characters).
        checkBox.SetCheckedSymbol(0x2611, "Arial");
        checkBox.SetUncheckedSymbol(0x2610, "Arial");

        // Insert the content control at the current cursor position.
        builder.InsertNode(checkBox);

        // Add placeholder text inside the content control.
        builder.Writeln("Please accept the terms and conditions.");

        // Save the document to a file. The format is inferred from the extension.
        doc.Save("ContentControlExample.docx");
    }
}
