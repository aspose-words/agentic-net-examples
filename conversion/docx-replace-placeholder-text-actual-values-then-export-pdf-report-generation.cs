using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace ReportGeneration
{
    class Program
    {
        static void Main()
        {
            // Load the template DOCX file.
            // This uses the Document(string) constructor – the approved "load" rule.
            Document doc = new Document("Template.docx");

            // Replace placeholder tags with actual values.
            // The Range.Replace method is the approved way to perform find‑and‑replace.
            doc.Range.Replace("_FullName_", "John Doe");
            doc.Range.Replace("_Company_", "Acme Corp");
            doc.Range.Replace("_Date_", DateTime.Today.ToString("MMMM d, yyyy"));

            // Save the populated document as PDF.
            // This uses the Document.Save(string, SaveFormat) overload – the approved "save" rule.
            doc.Save("Report.pdf", SaveFormat.Pdf);
        }
    }
}
