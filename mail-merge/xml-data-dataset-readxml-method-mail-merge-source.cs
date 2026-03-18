using System;
using System.Data;
using Aspose.Words;

namespace AsposeWordsMailMergeExample
{
    class Program
    {
        static void Main()
        {
            // Path to the folder that contains the template and XML files.
            string dataDir = @"C:\MyProject\Resources\";

            // Load the Word template that contains mail‑merge fields and regions.
            Document doc = new Document(dataDir + "Template.docx");

            // Create a DataSet and populate it from the XML file.
            // The ReadXml method parses the XML and builds tables that match the XML structure.
            DataSet dataSet = new DataSet();
            dataSet.ReadXml(dataDir + "Data.xml");

            // Perform mail merge using the DataSet as the source.
            // This uses the ExecuteWithRegions overload that accepts a DataSet.
            doc.MailMerge.ExecuteWithRegions(dataSet);

            // Save the merged document.
            doc.Save(dataDir + "MergedOutput.docx");
        }
    }
}
