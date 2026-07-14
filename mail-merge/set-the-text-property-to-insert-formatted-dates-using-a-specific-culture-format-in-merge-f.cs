using System;
using System.Data;
using System.Globalization;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Fields;

namespace AsposeWordsMailMergeExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set the builder's font locale to German – this will be the locale stored in the field code.
            builder.Font.LocaleId = new CultureInfo("de-DE").LCID;

            // Insert two MERGEFIELDs with a date format switch.
            builder.InsertField("MERGEFIELD Date1 \\@ \"dddd, d MMMM yyyy\"");
            builder.Write(" - ");
            builder.InsertField("MERGEFIELD Date2 \\@ \"dddd, d MMMM yyyy\"");

            // Preserve the original thread culture.
            CultureInfo originalCulture = Thread.CurrentThread.CurrentCulture;

            // First merge: use the current thread's culture (English US) for formatting.
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            doc.MailMerge.Execute(new[] { "Date1" }, new object[] { new DateTime(2020, 1, 1) });

            // Second merge: tell Aspose.Words to obtain the culture from the field code (German).
            doc.FieldOptions.FieldUpdateCultureSource = FieldUpdateCultureSource.FieldCode;
            doc.MailMerge.Execute(new[] { "Date2" }, new object[] { new DateTime(2020, 1, 1) });

            // Output the merged result to the console.
            Console.WriteLine(doc.Range.Text.Trim());

            // Save the document (optional – demonstrates the save rule).
            doc.Save("MergedDates.docx");

            // Restore the original thread culture.
            Thread.CurrentThread.CurrentCulture = originalCulture;
        }
    }
}
