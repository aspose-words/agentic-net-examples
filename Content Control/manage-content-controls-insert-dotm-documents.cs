using System;
using Aspose.Words;
using Aspose.Words.Markup;

namespace AsposeWordsDemo
{
    class Program
    {
        static void Main()
        {
            // Load the DOTM template that we want to insert.
            // The file must exist at the specified path.
            Document dotmTemplate = new Document(@"C:\Templates\MyTemplate.dotm");

            // Create a new blank document that will receive the content.
            Document destination = new Document();

            // Initialize a DocumentBuilder for the destination document.
            DocumentBuilder builder = new DocumentBuilder(destination);

            // Write some text before the content control.
            builder.Writeln("Document start.");

            // Insert a plain‑text content control (StructuredDocumentTag) into the document.
            // This will act as the placeholder where the DOTM content will be placed.
            StructuredDocumentTag contentControl = builder.InsertStructuredDocumentTag(SdtType.PlainText);

            // Write some text after the content control (optional, just for demonstration).
            builder.Writeln("Document end.");

            // Move the builder's cursor to the start of the content control.
            builder.MoveTo(contentControl);

            // Insert the DOTM document at the current cursor position.
            // KeepSourceFormatting preserves the original formatting of the template.
            builder.InsertDocument(dotmTemplate, ImportFormatMode.KeepSourceFormatting);

            // Save the resulting document.
            destination.Save(@"C:\Output\Result.docx");
        }
    }
}
