using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create sample documents that will be inserted for each data row.
        // -----------------------------------------------------------------
        string docPath1 = Path.Combine(outputDir, "Sample1.docx");
        string docPath2 = Path.Combine(outputDir, "Sample2.docx");

        CreateSampleDocument(docPath1, "Content of the first inserted document.");
        CreateSampleDocument(docPath2, "Content of the second inserted document.");

        // ---------------------------------------------------------------
        // 2. Build a mail‑merge template that contains a MERGEFIELD.
        //    The field value will be the file name of the document to insert.
        // ---------------------------------------------------------------
        string templatePath = Path.Combine(outputDir, "Template.docx");
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("=== Begin of merged document ===");
        // The field name can be anything; we use "DocPath".
        builder.InsertField("MERGEFIELD DocPath");
        builder.Writeln(); // Ensure a line break after each inserted document.
        builder.Writeln("=== End of merged document ===");

        template.Save(templatePath, SaveFormat.Docx);

        // ---------------------------------------------------------------
        // 3. Prepare a data source (DataTable) with a row for each document.
        // ---------------------------------------------------------------
        DataTable data = new DataTable("Docs");
        data.Columns.Add("DocPath", typeof(string));
        data.Rows.Add(docPath1);
        data.Rows.Add(docPath2);

        // ---------------------------------------------------------------
        // 4. Subscribe to the FieldMergingCallback.
        //    When the "DocPath" field is encountered we load the referenced
        //    document and insert its content at the field location.
        // ---------------------------------------------------------------
        template.MailMerge.FieldMergingCallback = new InsertDocumentCallback();

        // ---------------------------------------------------------------
        // 5. Execute the mail merge. For each row the callback will insert
        //    the corresponding document.
        // ---------------------------------------------------------------
        template.MailMerge.Execute(data);

        // ---------------------------------------------------------------
        // 6. Save the combined result as PDF.
        // ---------------------------------------------------------------
        string pdfPath = Path.Combine(outputDir, "CombinedResult.pdf");
        template.Save(pdfPath, SaveFormat.Pdf);

        // Simple validation that the PDF was created.
        if (File.Exists(pdfPath))
        {
            Console.WriteLine($"Combined PDF generated successfully at: {pdfPath}");
        }
        else
        {
            throw new InvalidOperationException("Failed to generate the combined PDF file.");
        }
    }

    // Helper method to create a minimal DOCX with a single paragraph of text.
    private static void CreateSampleDocument(string filePath, string text)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln(text);
        doc.Save(filePath, SaveFormat.Docx);
    }

    // Callback that replaces the MERGEFIELD value with the contents of the
    // document whose file name is stored in the field.
    private class InsertDocumentCallback : IFieldMergingCallback
    {
        void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
        {
            // We only act on the "DocPath" field.
            if (!args.DocumentFieldName.Equals("DocPath", StringComparison.OrdinalIgnoreCase))
                return;

            string fileName = args.FieldValue?.ToString();

            if (string.IsNullOrEmpty(fileName) || !File.Exists(fileName))
            {
                // Insert a placeholder text if the file is missing.
                args.Text = "[Missing document]";
                return;
            }

            // Load the document to be inserted.
            Document insertDoc = new Document(fileName);

            // Move the builder to the merge field location.
            DocumentBuilder builder = new DocumentBuilder(args.Document);
            builder.MoveToMergeField(args.DocumentFieldName);

            // Insert the whole document at the field position.
            builder.InsertDocument(insertDoc, ImportFormatMode.KeepSourceFormatting);

            // Suppress the default text insertion for this field.
            args.Text = string.Empty;
        }

        void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
        {
            // No image handling required for this example.
        }
    }
}
