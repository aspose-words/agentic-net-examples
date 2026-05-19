using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class Program
{
    public static void Main()
    {
        // Paths for temporary files.
        const string templatePath = "Template.docx";
        const string insertDocPath = "InsertDoc.docx";
        const string outputPdfPath = "MergedResult.pdf";

        // -----------------------------------------------------------------
        // 1. Create a simple template document with a merge field that will be
        //    replaced by the inserted DOCX.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder tmplBuilder = new DocumentBuilder(template);
        tmplBuilder.Writeln("Record ID: ");
        tmplBuilder.InsertField("MERGEFIELD RecordID");
        tmplBuilder.Writeln();
        tmplBuilder.InsertField("MERGEFIELD InsertDoc"); // Placeholder for the DOCX.
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Create the DOCX that will be inserted for each record.
        // -----------------------------------------------------------------
        Document insertDoc = new Document();
        DocumentBuilder insertBuilder = new DocumentBuilder(insertDoc);
        insertBuilder.Writeln("=== Inserted content for a record ===");
        insertBuilder.Writeln("This text comes from the external DOCX.");
        insertDoc.Save(insertDocPath);

        // -----------------------------------------------------------------
        // 3. Prepare a data source with two records.
        // -----------------------------------------------------------------
        DataTable data = new DataTable("Data");
        data.Columns.Add("RecordID", typeof(string));
        data.Rows.Add("1");
        data.Rows.Add("2");

        // -----------------------------------------------------------------
        // 4. Set up a field merging callback that inserts the DOCX at the
        //    location of the "InsertDoc" merge field.
        // -----------------------------------------------------------------
        template.MailMerge.FieldMergingCallback = new InsertDocCallback(insertDocPath);

        // -----------------------------------------------------------------
        // 5. Execute the mail merge.
        // -----------------------------------------------------------------
        template.MailMerge.Execute(data);

        // -----------------------------------------------------------------
        // 6. Save the merged document as PDF.
        // -----------------------------------------------------------------
        template.Save(outputPdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // 7. Simple validation – ensure the PDF was created.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPdfPath))
            throw new InvalidOperationException("The PDF output was not created.");

        // Cleanup temporary files (optional).
        File.Delete(templatePath);
        File.Delete(insertDocPath);
    }

    // -----------------------------------------------------------------
    // Callback implementation.
    // -----------------------------------------------------------------
    private class InsertDocCallback : IFieldMergingCallback
    {
        private readonly string _docPath;

        public InsertDocCallback(string docPath)
        {
            _docPath = docPath;
        }

        public void FieldMerging(FieldMergingArgs args)
        {
            // Only act on the specific merge field.
            if (args.DocumentFieldName == "InsertDoc")
            {
                // Load the source DOCX that will be inserted.
                Document srcDoc = new Document(_docPath);

                // Move the builder to the merge field location.
                DocumentBuilder builder = new DocumentBuilder(args.Document);
                builder.MoveToMergeField(args.DocumentFieldName);

                // Insert the document while keeping its original formatting.
                // This also brings over language settings, so an explicit
                // ImportLanguageInfo call is not required.
                builder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

                // Prevent the default field text from being inserted.
                args.Text = string.Empty;
            }
        }

        public void ImageFieldMerging(ImageFieldMergingArgs args)
        {
            // No image handling required for this example.
        }
    }
}
