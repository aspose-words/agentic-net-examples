using System;
using Aspose.Words;
using Aspose.Words.MailMerging;
using Aspose.Words.Tables;

class MailMergeWithFooter
{
    static void Main()
    {
        // Load the existing template document.
        Document doc = new Document("Template.docx");

        // Create a DocumentBuilder for editing the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the builder cursor to the primary footer of the first section.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);

        // Insert static footer text. You can add line breaks or formatting as needed.
        builder.Write("Confidential – Company Name");
        builder.Writeln(); // optional extra line break

        // Prepare mail merge data.
        string[] fieldNames = { "FirstName", "LastName", "Address" };
        object[] fieldValues = { "John", "Doe", "123 Main St." };

        // Execute the mail merge using the prepared data.
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Save the merged document to a new file.
        doc.Save("Result.docx");
    }
}
