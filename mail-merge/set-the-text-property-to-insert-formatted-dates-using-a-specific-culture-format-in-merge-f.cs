using System;
using System.Globalization;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class Program
{
    public static void Main()
    {
        // Create a new document and add a merge field for the date.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Report Date:");
        builder.InsertField("MERGEFIELD Date \\* MERGEFORMAT");

        // Assign a custom callback to format the date during mail merge.
        doc.MailMerge.FieldMergingCallback = new DateFormattingCallback();

        // Execute mail merge with the current date.
        doc.MailMerge.Execute(
            new string[] { "Date" },
            new object[] { DateTime.Now });

        // Save the resulting document.
        doc.Save("FormattedDateOutput.docx");
    }

    // Custom callback to format date fields using a specific culture.
    private class DateFormattingCallback : IFieldMergingCallback
    {
        public void FieldMerging(FieldMergingArgs e)
        {
            if (e.FieldName == "Date" && e.FieldValue is DateTime dt)
            {
                // Format the date as "dd MMMM yyyy" in French culture.
                e.Text = dt.ToString("dd MMMM yyyy", new CultureInfo("fr-FR"));
            }
        }

        // Required by the interface but not used for image fields.
        public void ImageFieldMerging(ImageFieldMergingArgs args) { }
    }
}
