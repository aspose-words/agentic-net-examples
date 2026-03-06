using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the Word template that contains MERGEFIELD tags.
        Document doc = new Document("Template.docx");

        // Load XML data into a DataSet. The XML file should have a structure that matches the merge fields.
        DataSet dataSet = new DataSet();
        dataSet.ReadXml("Data.xml");

        // Perform mail merge using the DataSet. This method works with mail‑merge regions; if the template
        // does not contain regions the whole document will be merged once.
        doc.MailMerge.ExecuteWithRegions(dataSet);

        // Prepare XPS save options (default options are sufficient for this example).
        XpsSaveOptions xpsOptions = new XpsSaveOptions();

        // Save the merged document as XPS.
        doc.Save("Result.xps", xpsOptions);
    }
}
