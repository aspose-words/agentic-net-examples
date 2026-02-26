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

        // Add introductory text.
        builder.Writeln("Please fill in the following form:");

        // Create a plain‑text content control (StructuredDocumentTag).
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Block);
        sdt.Title = "Name";
        sdt.Tag = "NameTag";
        sdt.PlaceholderName = "Enter your name";

        // Insert the content control at the current cursor position.
        builder.InsertNode(sdt);

        // Add a default run inside the content control.
        sdt.AppendChild(new Run(doc, "John Doe"));

        // Protect the document so that only form fields (content controls) can be edited.
        doc.Protect(ProtectionType.AllowOnlyFormFields);

        // Save the protected document as a DOCX file.
        doc.Save("ProtectedDocument.docx");
    }
}
