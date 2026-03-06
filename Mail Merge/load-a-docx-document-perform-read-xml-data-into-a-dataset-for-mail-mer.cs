using System.Data;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the source DOCX template that contains MERGEFIELD fields.
        const string docPath = "Template.docx";

        // Path to the XML file that holds the mail‑merge data.
        const string xmlPath = "Data.xml";

        // Path where the merged result will be saved as XPS.
        const string outputPath = "Result.xps";

        // Load the Word document.
        Document doc = new Document(docPath);

        // Load the XML data into a DataSet.
        DataSet dataSet = new DataSet();
        dataSet.ReadXml(xmlPath);

        // Perform mail merge using the first table of the DataSet.
        // The table's column names must match the MERGEFIELD names in the template.
        if (dataSet.Tables.Count > 0)
        {
            doc.MailMerge.Execute(dataSet.Tables[0]);
        }

        // Save the merged document in XPS format.
        doc.Save(outputPath, SaveFormat.Xps);
    }
}
