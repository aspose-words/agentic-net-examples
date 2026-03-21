using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace ReportGeneration
{
    class Program
    {
        static void Main()
        {
            // Create a new document with placeholder text.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("_FullName_");
            builder.Writeln("_Company_");
            builder.Writeln("_Date_");

            // Replace placeholder tags with actual values.
            doc.Range.Replace("_FullName_", "John Doe");
            doc.Range.Replace("_Company_", "Acme Corp");
            doc.Range.Replace("_Date_", DateTime.Today.ToString("MMMM d, yyyy"));

            // Save the populated document as PDF.
            doc.Save("Report.pdf", SaveFormat.Pdf);
        }
    }
}
