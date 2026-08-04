using System;
using System.Data;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class Program
{
    public static void Main()
    {
        // Prepare a folder for temporary files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string templatePath = Path.Combine(artifactsDir, "Template.docx");
        string insertDocPath = Path.Combine(artifactsDir, "Insert.docx");
        string outputPdfPath = Path.Combine(artifactsDir, "Result.pdf");

        // -----------------------------------------------------------------
        // 1. Create a simple template document with merge fields.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder tmplBuilder = new DocumentBuilder(template);
        tmplBuilder.Writeln("Dear <<Name>>,");
        tmplBuilder.InsertField("MERGEFIELD Name");
        tmplBuilder.Writeln();
        tmplBuilder.InsertField("MERGEFIELD InsertDoc"); // placeholder for the DOCX to insert.
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Create a DOCX that will be inserted during mail merge.
        // -----------------------------------------------------------------
        Document insertDoc = new Document();
        DocumentBuilder insertBuilder = new DocumentBuilder(insertDoc);
        insertBuilder.Writeln("=== Inserted Document Content ===");
        insertBuilder.Writeln("This text comes from the inserted DOCX file.");
        // Set a language for demonstration (English - United States).
        // The Font class uses LocaleId (LCID) to specify language.
        insertBuilder.Font.LocaleId = CultureInfo.GetCultureInfo("en-US").LCID; // 1033
        insertDoc.Save(insertDocPath);

        // -----------------------------------------------------------------
        // 3. Prepare a data source with a record that references the DOCX.
        // -----------------------------------------------------------------
        DataTable data = new DataTable("Data");
        data.Columns.Add("Name", typeof(string));
        data.Columns.Add("InsertDoc", typeof(string));
        data.Rows.Add("John Doe", insertDocPath); // Single record.

        // -----------------------------------------------------------------
        // 4. Perform mail merge with a custom field merging callback that
        //    inserts the DOCX at the merge field location.
        // -----------------------------------------------------------------
        Document src = new Document(templatePath);
        src.MailMerge.FieldMergingCallback = new InsertDocCallback();
        src.MailMerge.Execute(data);

        // -----------------------------------------------------------------
        // 5. Save the final document as PDF.
        // -----------------------------------------------------------------
        src.Save(outputPdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // 6. Validate that the PDF was created.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPdfPath))
            throw new InvalidOperationException("The PDF output file was not created.");
    }

    // Custom callback to handle insertion of a DOCX during mail merge.
    private class InsertDocCallback : IFieldMergingCallback
    {
        public void FieldMerging(FieldMergingArgs args)
        {
            // Only handle the specific merge field used for document insertion.
            if (args.DocumentFieldName.Equals("InsertDoc", StringComparison.OrdinalIgnoreCase))
            {
                // The field value is expected to be the full path to the DOCX file.
                string docPath = args.FieldValue?.ToString();
                if (!string.IsNullOrEmpty(docPath) && File.Exists(docPath))
                {
                    // Load the source document to be inserted.
                    Document docToInsert = new Document(docPath);

                    // Move the builder to the merge field location.
                    DocumentBuilder builder = new DocumentBuilder(args.Document);
                    builder.MoveToMergeField(args.DocumentFieldName);

                    // Insert the document while keeping its original formatting (including language).
                    builder.InsertDocument(docToInsert, ImportFormatMode.KeepSourceFormatting);
                }

                // Suppress the default text that would otherwise be inserted.
                args.Text = string.Empty;
            }
        }

        public void ImageFieldMerging(ImageFieldMergingArgs args)
        {
            // No image handling required for this example.
        }
    }
}
