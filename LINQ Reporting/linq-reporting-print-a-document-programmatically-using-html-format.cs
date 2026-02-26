using System;
using System.Diagnostics;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // 1. Create a new blank Word document.
        Document doc = new Document();                     // Document()
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello from Aspose.Words!");      // add some content

        // 2. Populate the document using LINQ Reporting (optional demonstration).
        var data = new { Title = "Sample Report", Date = DateTime.Now };
        ReportingEngine engine = new ReportingEngine();    // ReportingEngine()
        engine.BuildReport(doc, data, "model");            // BuildReport(Document, object, string)

        // 3. Save the document as HTML to a temporary file.
        string htmlPath = Path.Combine(Path.GetTempPath(), "report.html");
        doc.Save(htmlPath, SaveFormat.Html);               // Save(string, SaveFormat)

        // 4. Print the generated HTML file using the default system printer.
        //    The "print" verb tells the OS to send the file to the printer without showing UI.
        ProcessStartInfo psi = new ProcessStartInfo
        {
            FileName = htmlPath,
            Verb = "print",
            CreateNoWindow = true,
            WindowStyle = ProcessWindowStyle.Hidden
        };

        // Start the process and wait for it to finish.
        using (Process? proc = Process.Start(psi))
        {
            proc?.WaitForExit();
        }
    }
}
