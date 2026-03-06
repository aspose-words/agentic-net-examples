using System;
using System.Data;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the mail‑merge template document (must contain MERGEFIELD tags or regions).
        string templatePath = "Template.docx";

        // Path to the XML file that holds the data for the merge.
        string xmlDataPath = "Data.xml";

        // Path where the merged document will be saved.
        string outputPath = "Merged.docx";

        // Load the template document.
        Document doc = new Document(templatePath);

        // Load the XML data into a DataSet. Aspose.Words can use a DataSet for mail merge.
        DataSet dataSet = new DataSet();
        dataSet.ReadXml(xmlDataPath);

        // Perform the mail merge using the DataSet. This works with both simple fields
        // and mail‑merge regions defined in the template.
        doc.MailMerge.ExecuteWithRegions(dataSet);

        // Save the merged document.
        doc.Save(outputPath);
    }
}
