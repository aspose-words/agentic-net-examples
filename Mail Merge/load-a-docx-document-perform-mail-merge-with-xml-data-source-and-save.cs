using Aspose.Words;
using System.Data;

class MailMergeXmlToXps
{
    static void Main()
    {
        // Input and output file paths
        string templatePath = "Template.docx";
        string xmlDataPath = "Data.xml";
        string resultPath = "Result.xps";

        // Load the DOCX template
        Document doc = new Document(templatePath);

        // Load XML data into a DataSet
        DataSet dataSet = new DataSet();
        dataSet.ReadXml(xmlDataPath);

        // Execute mail merge using the first table from the DataSet
        if (dataSet.Tables.Count > 0)
        {
            doc.MailMerge.Execute(dataSet.Tables[0]);
        }

        // Save the merged document as XPS
        doc.Save(resultPath, SaveFormat.Xps);
    }
}
