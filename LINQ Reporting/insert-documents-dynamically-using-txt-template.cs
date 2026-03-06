using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsDynamicInsert
{
    class Program
    {
        static void Main()
        {
            // Paths to the files – adjust as needed.
            string txtTemplatePath = @"C:\Templates\ReportTemplate.txt";
            string documentToInsertPath = @"C:\Documents\Appendix.docx";
            string outputPath = @"C:\Output\GeneratedReport.docx";

            // Load the TXT template into a Document.
            // Aspose.Words automatically detects the format from the file extension.
            Document doc = new Document(txtTemplatePath);

            // Create a DocumentBuilder attached to the loaded document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Example placeholders in the TXT template:
            // {{Title}}, {{Date}}, {{Author}}
            // Replace them with actual values using the Range.Replace method.
            doc.Range.Replace("{{Title}}", "Quarterly Financial Report");
            doc.Range.Replace("{{Date}}", DateTime.Now.ToString("MMMM dd, yyyy"));
            doc.Range.Replace("{{Author}}", "John Doe");

            // Move the cursor to the end of the document where we want to insert another document.
            builder.MoveToDocumentEnd();

            // Insert a page break before appending the new document for visual separation.
            builder.InsertBreak(BreakType.PageBreak);

            // Load the document that will be inserted.
            Document docToInsert = new Document(documentToInsertPath);

            // Insert the document using the KeepSourceFormatting mode to preserve its original styles.
            builder.InsertDocument(docToInsert, ImportFormatMode.KeepSourceFormatting);

            // Save the final document.
            doc.Save(outputPath, SaveFormat.Docx);
        }
    }
}
