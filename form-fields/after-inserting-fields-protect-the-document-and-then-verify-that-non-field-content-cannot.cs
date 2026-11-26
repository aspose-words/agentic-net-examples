using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert non‑field content.
        builder.Writeln("This is non‑field content.");

        // Insert a text input form field.
        builder.InsertTextInput("TextField", TextFormFieldType.Regular, "", "Default", 0);

        // Insert a checkbox form field.
        builder.InsertCheckBox("CheckBox", false, 0);

        // Protect the document for read‑only editing (forms cannot be edited while protected).
        // The password is required for later unprotecting.
        doc.Protect(ProtectionType.ReadOnly, "pwd");

        // Attempt to edit the non‑field paragraph (should be blocked while protected).
        try
        {
            Paragraph para = doc.FirstSection.Body.Paragraphs[0];
            if (para != null && para.Runs.Count > 0)
            {
                para.Runs[0].Text = "Attempted edit";
                Console.WriteLine("Non‑field edit succeeded unexpectedly (the document may not enforce protection).");
            }
            else
            {
                Console.WriteLine("Paragraph or run not found.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Non‑field edit blocked: {ex.Message}");
        }

        // Unprotect the document to allow form field editing.
        doc.Unprotect("pwd");

        // Edit the text form field (allowed after unprotecting).
        FormField textField = doc.Range.FormFields["TextField"];
        if (textField != null)
        {
            textField.Result = "NewValue";
            Console.WriteLine("Form field edited successfully.");
        }
        else
        {
            throw new InvalidOperationException("Form field 'TextField' not found.");
        }

        // Re‑protect the document (optional, to keep it protected after modifications).
        doc.Protect(ProtectionType.ReadOnly, "pwd");

        // Save the protected document.
        doc.Save("ProtectedDocument.docx");
        Console.WriteLine("Document saved as 'ProtectedDocument.docx'.");
    }
}
