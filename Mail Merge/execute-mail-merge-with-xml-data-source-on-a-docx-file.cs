using System;
using System.Data;
using Aspose.Words;

namespace MailMergeXmlExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source DOCX file that contains MERGEFIELD tags.
            const string sourceDocPath = @"C:\Docs\Template.docx";

            // Path to the XML file that holds the mail‑merge data.
            const string xmlDataPath = @"C:\Docs\Data.xml";

            // Path where the merged document will be saved.
            const string outputDocPath = @"C:\Docs\MergedResult.docx";

            // Load the template document.
            Document doc = new Document(sourceDocPath);

            // Load the XML data into a DataSet.
            DataSet dataSet = new DataSet();
            dataSet.ReadXml(xmlDataPath);

            // Assume the XML contains a single table; use the first DataTable as the data source.
            // If the XML defines multiple tables, select the appropriate one by name.
            DataTable table = dataSet.Tables[0];

            // Perform the mail merge using the DataTable.
            doc.MailMerge.Execute(table);

            // Save the merged document.
            doc.Save(outputDocPath);
        }
    }
}
