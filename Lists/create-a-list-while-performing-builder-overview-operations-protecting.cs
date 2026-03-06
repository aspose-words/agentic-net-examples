using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Lists; // Added for ListTemplate

class Program
{
    static void Main()
    {
        // Load an existing DOCX document.
        Document docOriginal = new Document("Original.docx");

        // -----------------------------------------------------------------
        // Builder Overview: create a bullet list using DocumentBuilder.
        // -----------------------------------------------------------------
        DocumentBuilder builder = new DocumentBuilder(docOriginal);
        // Ensure the document has a list style to use.
        builder.ListFormat.List = docOriginal.Lists.Add(ListTemplate.BulletDefault);
        // Add three list items.
        builder.Writeln("First item");
        builder.Writeln("Second item");
        builder.Writeln("Third item");
        // End the list formatting.
        builder.ListFormat.RemoveNumbers();

        // -----------------------------------------------------------------
        // Protect the document with a password (read‑only protection).
        // -----------------------------------------------------------------
        docOriginal.Protect(ProtectionType.ReadOnly, "protectPwd");

        // Save the protected document (no encryption).
        docOriginal.Save("Protected.docx");

        // -----------------------------------------------------------------
        // Save the same document with encryption (password protection on save).
        // -----------------------------------------------------------------
        OoxmlSaveOptions encryptOptions = new OoxmlSaveOptions(SaveFormat.Docx)
        {
            Password = "encryptPwd"
        };
        docOriginal.Save("Encrypted.docx", encryptOptions);

        // -----------------------------------------------------------------
        // Prepare a second document to compare with.
        // Clone the original (before protection) and make a simple edit.
        // -----------------------------------------------------------------
        Document docForComparison = (Document)docOriginal.Clone(true);
        // Remove protection from the clone so we can edit it.
        docForComparison.Unprotect();
        // Append a new paragraph.
        DocumentBuilder editBuilder = new DocumentBuilder(docForComparison);
        editBuilder.Writeln("Additional paragraph added for comparison.");

        // -----------------------------------------------------------------
        // Compare the original (protected) document with the edited one.
        // The comparison will add revision marks to the original document.
        // -----------------------------------------------------------------
        docOriginal.Compare(docForComparison, "Comparer", DateTime.Now);

        // Save the document that now contains revision marks.
        docOriginal.Save("Compared.docx");
    }
}
