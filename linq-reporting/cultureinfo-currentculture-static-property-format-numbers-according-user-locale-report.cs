using System;
using System.Globalization;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Fields;

class ReportGenerator
{
    static void Main()
    {
        // Use the culture of the current user (thread) for number formatting.
        CultureInfo userCulture = CultureInfo.CurrentCulture;
        Thread.CurrentThread.CurrentCulture = userCulture;

        // Create a new document and a builder to add content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Ensure fields use the thread's culture (default behavior, set explicitly for clarity).
        doc.FieldOptions.UseInvariantCultureNumberFormat = false;
        doc.FieldOptions.FieldUpdateCultureSource = FieldUpdateCultureSource.CurrentThread;

        // Insert a numeric field. The format string "#,##0.00" will be interpreted using the current culture.
        // For example, in en-US it becomes "12,345.67", while in de-DE it becomes "12.345,67".
        builder.InsertField("= 12345.67 \\# #,##0.00");

        // Calculate and format the field result.
        doc.UpdateFields();

        // Save the document.
        doc.Save("Report.docx");
    }
}
