using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.MailMerging;

namespace MailMergeExample
{
    class Program
    {
        static void Main()
        {
            // Path to the template document that contains MERGEFIELDs and region tags.
            // Example template should have fields like <<TableStart:Employees>> ... <<TableEnd:Employees>>
            string templatePath = @"C:\Templates\EmployeeTemplate.docx";

            // Path to the XML file that will be loaded into a DataSet.
            // The XML should contain tables whose names match the mail‑merge region names.
            string xmlDataPath = @"C:\Data\Employees.xml";

            // Path where the merged document will be saved.
            string outputPath = @"C:\Output\EmployeeReport.docx";

            // Load the template document.
            Document doc = new Document(templatePath);

            // Load XML data into a DataSet.
            DataSet dataSet = new DataSet();
            dataSet.ReadXml(xmlDataPath);

            // Perform mail merge using the DataSet.
            // This method will look for mail‑merge regions whose names correspond to the DataTable names.
            doc.MailMerge.ExecuteWithRegions(dataSet);

            // Save the merged document.
            doc.Save(outputPath);
        }
    }
}
