using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsInsertDynamic
{
    class Program
    {
        static void Main()
        {
            // Load the DOTM template that contains the placeholder where the document will be inserted.
            // The template can be a normal Word document with a .dotm extension.
            Document template = new Document(@"C:\Templates\MyTemplate.dotm");

            // Create a new document that will hold the final result.
            Document resultDoc = new Document();

            // Use DocumentBuilder to work with the result document.
            DocumentBuilder builder = new DocumentBuilder(resultDoc);

            // Optionally, copy styles from the template to the result document.
            // This ensures that any style referenced in the template exists in the result.
            resultDoc.CopyStylesFromTemplate(template);

            // Move the cursor to the end of the document (or to a bookmark if you have one).
            builder.MoveToDocumentEnd();

            // Load the source document that you want to insert dynamically.
            Document sourceDoc = new Document(@"C:\Data\SourceDocument.docx");

            // Configure import options to enable the "sourceStyles" switch.
            // Setting SmartStyleBehavior to true makes Aspose.Words resolve style clashes
            // by converting source styles with the same name into direct formatting.
            ImportFormatOptions importOptions = new ImportFormatOptions
            {
                SmartStyleBehavior = true
            };

            // Insert the source document into the result document using KeepSourceFormatting.
            // This preserves the original formatting of the source while applying the smart style behavior.
            builder.InsertDocument(sourceDoc, ImportFormatMode.KeepSourceFormatting, importOptions);

            // If you need to populate the template with data, you can use ReportingEngine here.
            // Example (optional):
            // ReportingEngine engine = new ReportingEngine();
            // engine.BuildReport(template, dataSourceObject, "data");

            // Save the final document.
            resultDoc.Save(@"C:\Output\CombinedResult.docx");
        }
    }
}
