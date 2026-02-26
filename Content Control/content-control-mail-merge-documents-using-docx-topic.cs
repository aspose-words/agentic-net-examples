using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.BuildingBlocks;
using Aspose.Words.Markup;

namespace AsposeWordsContentControlMailMerge
{
    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a new blank document.
            // -----------------------------------------------------------------
            Document doc = new Document();                     // create
            DocumentBuilder builder = new DocumentBuilder(doc);

            // -----------------------------------------------------------------
            // 2. Insert a Structured Document Tag (Content Control) that will
            //    hold the merge field placeholder.
            // -----------------------------------------------------------------
            // The tag type can be PlainText; inside we place a placeholder that
            // Aspose.Words will replace when UseNonMergeFields is enabled.
            StructuredDocumentTag sdt = builder.InsertStructuredDocumentTag(SdtType.PlainText);
            sdt.Title = "FullName";
            sdt.Tag = "FullName";
            // Placeholder text using the mustache syntax.
            sdt.AppendChild(new Run(doc, "{{FullName}}"));

            // Add a paragraph break after the content control.
            builder.Writeln();

            // Insert another content control for the address.
            StructuredDocumentTag addressSdt = builder.InsertStructuredDocumentTag(SdtType.PlainText);
            addressSdt.Title = "Address";
            addressSdt.Tag = "Address";
            addressSdt.AppendChild(new Run(doc, "{{Address}}"));
            builder.Writeln();

            // -----------------------------------------------------------------
            // 3. Configure the MailMerge engine to treat the mustache tags as
            //    merge fields.
            // -----------------------------------------------------------------
            doc.MailMerge.UseNonMergeFields = true;   // enable non‑merge field processing

            // -----------------------------------------------------------------
            // 4. Perform the mail merge using an array of field names and values.
            // -----------------------------------------------------------------
            string[] fieldNames = { "FullName", "Address" };
            object[] fieldValues = { "John Doe", "123 Main St, Anytown" };
            doc.MailMerge.Execute(fieldNames, fieldValues);

            // -----------------------------------------------------------------
            // 5. Save the merged document to a DOCX file.
            // -----------------------------------------------------------------
            string outputPath = Path.Combine(Environment.CurrentDirectory, "MergedDocument.docx");
            doc.Save(outputPath);                     // save
        }
    }
}
