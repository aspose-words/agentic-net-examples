using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Saving;

class MailMergeToXps
{
    static void Main()
    {
        // Path to the source DOCX file that contains MERGEFIELD tags.
        const string docxPath = @"C:\Input\Template.docx";

        // Path to the XML file that holds the data for the mail merge.
        const string xmlPath = @"C:\Input\Data.xml";

        // Path where the resulting XPS file will be saved.
        const string xpsPath = @"C:\Output\Result.xps";

        // Load the Word document.
        Document doc = new Document(docxPath);

        // Load XML data into a DataSet.
        DataSet dataSet = new DataSet();
        dataSet.ReadXml(xmlPath);

        // Assume the XML contains at least one table; use the first table for the merge.
        // If the document uses mail‑merge regions, use ExecuteWithRegions(dataSet) instead.
        if (dataSet.Tables.Count > 0)
        {
            DataTable table = dataSet.Tables[0];
            doc.MailMerge.Execute(table);
        }
        else
        {
            throw new InvalidOperationException("No tables were found in the XML data source.");
        }

        // Save the merged document as XPS.
        doc.Save(xpsPath, SaveFormat.Xps);
    }
}
