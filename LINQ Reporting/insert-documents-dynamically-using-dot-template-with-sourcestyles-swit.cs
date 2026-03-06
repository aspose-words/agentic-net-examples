using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the DOT template that contains the base layout and styles.
        Document template = new Document("Template.dotx");

        // Enable automatic style updating so that any changes in the template
        // are applied to the document when it is opened in Word.
        template.AutomaticallyUpdateStyles = true;

        // List of source documents that will be inserted into the template.
        List<string> sourceFiles = new List<string>
        {
            "Part1.docx",
            "Part2.docx",
            "Part3.docx"
        };

        // Create a DocumentBuilder for the template.
        DocumentBuilder builder = new DocumentBuilder(template);

        // Position the cursor where the inserts should start.
        builder.MoveToDocumentEnd();

        // Configure import options to resolve style name clashes by expanding the
        // source style into direct formatting (SmartStyleBehavior).
        ImportFormatOptions importOptions = new ImportFormatOptions
        {
            SmartStyleBehavior = true
        };

        // Insert each source document while preserving its original formatting.
        foreach (string filePath in sourceFiles)
        {
            Document srcDoc = new Document(filePath);

            // KeepSourceFormatting ensures the source document looks the same after insertion.
            builder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting, importOptions);

            // Insert a page break between documents (optional).
            builder.InsertBreak(BreakType.PageBreak);
        }

        // Save the final merged document.
        template.Save("MergedResult.docx");
    }
}
