using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a field that outputs the document's creation date/time
        // formatted as ISO‑8601 (yyyy‑MM‑ddTHH:mm:ss).
        // The expression {=CreatedDate:format} is equivalent to the field code:
        //   = CreatedDate \@ "format"
        builder.InsertField("{=CreatedDate:yyyy-MM-ddTHH:mm:ss}");

        // Evaluate the field.
        doc.UpdateFields();

        // Save the document.
        doc.Save("CreatedDateISO8601.docx");
    }
}
