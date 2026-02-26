using System;
using System.Diagnostics;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // 1. Create a new blank Word document.
        Document doc = new Document();

        // 2. Add some content using DocumentBuilder.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello world! This document will be saved as PDF and printed to a PDF printer.");

        // 3. Save the document as PDF. The .pdf extension automatically selects the PDF format.
        string pdfPath = "Report.pdf";
        doc.Save(pdfPath, SaveFormat.Pdf);

        // 4. Print the generated PDF using the default PDF printer (e.g., "Microsoft Print to PDF").
        //    In .NET Core/.NET 5+ Aspose.Words does not expose a Document.Print method, so we launch the PDF
        //    file with the "print" verb. The system will route the job to the default printer, which should be a
        //    PDF printer if you want a PDF output.
        var psi = new ProcessStartInfo
        {
            FileName = pdfPath,
            Verb = "print",
            CreateNoWindow = true,
            UseShellExecute = true
        };
        try
        {
            Process.Start(psi);
            Console.WriteLine($"Print job sent for '{pdfPath}'. Ensure a PDF printer is set as the default printer.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to print the PDF: {ex.Message}");
        }
    }
}
