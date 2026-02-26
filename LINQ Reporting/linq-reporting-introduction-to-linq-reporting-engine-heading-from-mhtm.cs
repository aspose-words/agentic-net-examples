using System;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace LinqReportingExample
{
    class Program
    {
        static void Main()
        {
            // Load the source MHTML file. Aspose.Words automatically detects the format.
            Document doc = new Document("InputReport.mht");

            // Insert a heading at the beginning of the document.
            DocumentBuilder builder = new DocumentBuilder(doc);
            // Move the cursor to the start of the document.
            builder.MoveToDocumentStart();
            // Apply Heading 1 style.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            // Write the heading text.
            builder.Writeln("LINQ Reporting Introduction to LINQ Reporting Engine");

            // Save the modified document as a DOCX file.
            doc.Save("OutputReport.docx");
        }
    }
}
