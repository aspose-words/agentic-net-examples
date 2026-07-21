using System;
using System.Globalization;
using Aspose.Words;
using Aspose.Words.Markup;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a simple source DOC file if it does not exist.
        const string sourcePath = "input.doc";
        if (!System.IO.File.Exists(sourcePath))
        {
            Document seed = new Document();
            DocumentBuilder seedBuilder = new DocumentBuilder(seed);
            seedBuilder.Writeln("Sample document for adding a date picker content control.");
            seed.Save(sourcePath);
        }

        // Step 2: Load the existing DOC file.
        Document doc = new Document(sourcePath);

        // Step 3: Create a date picker content control (StructuredDocumentTag of type Date).
        StructuredDocumentTag dateSdt = new StructuredDocumentTag(doc, SdtType.Date, MarkupLevel.Inline)
        {
            Title = "AppointmentDate",
            Tag = "appointment-date"
        };

        // Optional: configure display format, locale, and a default date.
        dateSdt.DateDisplayLocale = CultureInfo.GetCultureInfo("en-US").LCID;
        dateSdt.DateDisplayFormat = "dd MMMM, yyyy";
        dateSdt.DateStorageFormat = SdtDateStorageFormat.DateTime;
        dateSdt.FullDate = DateTime.Today;

        // Step 4: Insert the content control into the document.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln(); // Ensure we are on a new paragraph.
        builder.Write("Select a date: ");
        builder.InsertNode(dateSdt);

        // Step 5: Save the modified document as DOCX.
        const string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}
