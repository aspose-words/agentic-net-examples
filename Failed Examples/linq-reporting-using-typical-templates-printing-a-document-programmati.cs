// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using System.Drawing.Printing;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.LoadOptions = Aspose.Words.LoadOptions;
using Aspose.Words.Rendering;

namespace AsposeWordsLinqReportingDemo
{
    // Simple POCO class that will be used as a data source for the report.
    public class Employee
    {
        public string Name { get; set; }
        public string Position { get; set; }
        public decimal Salary { get; set; }
    }

    class Program
    {
        // Entry point.
        [STAThread]
        static void Main()
        {
            // 1. Load a WORDML template (XML based) from a file.
            //    The LoadFormat.WordML tells Aspose.Words to treat the file as WORDML.
            var loadOptions = new LoadOptions(LoadFormat.WordML);
            Document template = new Document("Template.xml", loadOptions);

            // 2. Prepare a collection of data that will be merged into the template.
            List<Employee> employees = new List<Employee>
            {
                new Employee { Name = "John Doe", Position = "Developer", Salary = 72000m },
                new Employee { Name = "Jane Smith", Position = "Designer", Salary = 68000m },
                new Employee { Name = "Bob Johnson", Position = "Manager", Salary = 95000m }
            };

            // 3. Use ReportingEngine to populate the template.
            //    The data source name "Employees" will be used inside the template
            //    with the syntax <<foreach [Employees]>><<[Name]>> - <<[Position]>> - <<[Salary]:currency>>><</foreach>>.
            ReportingEngine engine = new ReportingEngine();
            // Optional: allow missing members without throwing.
            engine.Options = ReportBuildOptions.AllowMissingMembers;
            engine.BuildReport(template, employees, "Employees");

            // 4. Save the generated report to a DOCX file.
            //    The Save method automatically determines the format from the extension.
            string reportPath = "GeneratedReport.docx";
            template.Save(reportPath);

            // 5. Update page layout before any rendering/printing operation.
            template.UpdatePageLayout();

            // 6. Print the document programmatically using the default printer.
            PrintDocument printDoc = new AsposeWordsPrintDocument(template);
            // Example: print silently without UI.
            printDoc.Print();

            // 7. Alternatively, let the user choose printer settings via a PrintDialog.
            using (PrintDialog dlg = new PrintDialog())
            {
                dlg.AllowSomePages = true;
                dlg.PrinterSettings.MinimumPage = 1;
                dlg.PrinterSettings.MaximumPage = template.PageCount;
                dlg.PrinterSettings.FromPage = 1;
                dlg.PrinterSettings.ToPage = template.PageCount;

                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    AsposeWordsPrintDocument awPrintDoc = new AsposeWordsPrintDocument(template);
                    awPrintDoc.PrinterSettings = dlg.PrinterSettings;
                    // Cache printer settings to speed up the first print call.
                    awPrintDoc.CachePrinterSettings();
                    awPrintDoc.Print();
                }
            }

            // 8. Demonstrate printing with a preview dialog.
            using (PrintPreviewDialog previewDlg = new PrintPreviewDialog())
            {
                AsposeWordsPrintDocument previewPrintDoc = new AsposeWordsPrintDocument(template);
                previewDlg.Document = previewPrintDoc;
                previewDlg.ShowDialog();
            }

            Console.WriteLine("Report generated, saved to '{0}' and printed.", reportPath);
        }
    }
}
