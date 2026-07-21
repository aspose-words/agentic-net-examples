using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

namespace AsposeWordsMailMergeInsertDoc
{
    public class Program
    {
        // Paths for the generated files.
        private const string OutputPdfPath = "MergedOutput.pdf";
        private const string TemplateDocPath = "Template.docx";
        private const string InsertDocPath = "InsertDocument.docx";

        public static void Main()
        {
            // Create sample source documents.
            CreateInsertDocument();
            CreateTemplateDocument();

            // Load the template that contains the MERGEFIELD.
            Document template = new Document(TemplateDocPath);

            // Register a callback that will replace the merge field with the content of InsertDocument.docx.
            template.MailMerge.FieldMergingCallback = new InsertDocFieldMergingCallback(InsertDocPath);

            // Dummy data source – the actual content is supplied by the callback.
            DataTable data = new DataTable("Data");
            data.Columns.Add("Doc", typeof(string));
            data.Rows.Add(string.Empty);
            data.Rows.Add(string.Empty); // Two records to demonstrate multiple inserts.

            // Perform the mail merge. The callback will insert the document at each field occurrence.
            template.MailMerge.Execute(data);

            // Save the merged result as PDF.
            template.Save(OutputPdfPath, SaveFormat.Pdf);

            // Verify that the PDF was created.
            if (!File.Exists(OutputPdfPath))
                throw new InvalidOperationException($"Failed to create the output file: {OutputPdfPath}");
        }

        // Creates a simple DOCX that will be inserted during mail merge.
        private static void CreateInsertDocument()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set language for the run using LocaleId (1033 = English - United States).
            builder.Font.LocaleId = 1033;
            builder.Writeln("This is the inserted document content.");

            doc.Save(InsertDocPath, SaveFormat.Docx);
        }

        // Creates a template DOCX containing a merge field where the document will be inserted.
        private static void CreateTemplateDocument()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("=== Start of Template ===");
            // Insert a merge field named "Doc". The callback will replace this field.
            builder.InsertField(" MERGEFIELD Doc ");
            builder.Writeln("=== End of Template ===");

            doc.Save(TemplateDocPath, SaveFormat.Docx);
        }

        // Custom callback that inserts a DOCX at the location of the merge field.
        private class InsertDocFieldMergingCallback : IFieldMergingCallback
        {
            private readonly string _sourceDocPath;

            public InsertDocFieldMergingCallback(string sourceDocPath)
            {
                _sourceDocPath = sourceDocPath;
            }

            public void FieldMerging(FieldMergingArgs args)
            {
                if (args.DocumentFieldName.Equals("Doc", StringComparison.OrdinalIgnoreCase))
                {
                    // Load the document to be inserted.
                    Document srcDoc = new Document(_sourceDocPath);

                    // Move the builder to the merge field location.
                    DocumentBuilder builder = new DocumentBuilder(args.Document);
                    builder.MoveToMergeField(args.DocumentFieldName);

                    // Insert the source document, preserving its formatting.
                    builder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

                    // Suppress the default field text.
                    args.Text = string.Empty;
                }
            }

            public void ImageFieldMerging(ImageFieldMergingArgs args)
            {
                // No image handling required for this example.
            }
        }
    }
}
