using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder to insert content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a normal paragraph (non‑form field content).
        builder.Writeln("This is a normal paragraph that should not be editable when the document is protected.");

        // Insert a text input form field.
        builder.Write("Enter your name: ");
        FormField textField = builder.InsertTextInput("NameField", TextFormFieldType.Regular, "", "John Doe", 50);
        builder.Writeln(); // Move to next line.

        // Insert another normal paragraph.
        builder.Writeln("Another normal paragraph that should remain read‑only under protection.");

        // Save the document before protection (optional, for inspection).
        string outputPath = Path.Combine(Environment.CurrentDirectory, "FormFieldsProtected.docx");
        doc.Save(outputPath);

        // Protect the document so that only form fields can be edited.
        doc.Protect(ProtectionType.AllowOnlyFormFields);

        // Verify that the document protection type is set correctly.
        if (doc.ProtectionType != ProtectionType.AllowOnlyFormFields)
            throw new InvalidOperationException("Document protection type was not set to AllowOnlyFormFields.");

        // Verify that the first section is protected for forms (non‑field content is read‑only in UI).
        if (!doc.Sections[0].ProtectedForForms)
            throw new InvalidOperationException("Section is not protected for forms.");

        // Attempt to modify non‑field content programmatically.
        // This simulates a user trying to edit a normal paragraph.
        Paragraph firstParagraph = doc.FirstSection.Body.Paragraphs[0];
        string originalText = firstParagraph.GetText();

        // Change the paragraph text.
        firstParagraph.Runs.Clear();
        firstParagraph.AppendChild(new Run(doc, "Edited paragraph text."));

        // After protection, Aspose.Words still allows programmatic changes,
        // but UI users cannot edit this text. We verify that the paragraph was changed.
        // If the text remains unchanged, it would indicate that protection prevented the edit.
        string modifiedText = firstParagraph.GetText();

        bool editSucceeded = !originalText.Equals(modifiedText, StringComparison.Ordinal);
        Console.WriteLine(editSucceeded
            ? "Programmatic edit succeeded (protection does not block code changes)."
            : "Edit was blocked by protection.");

        // Save the protected document.
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
