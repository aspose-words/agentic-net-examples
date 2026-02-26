using System;
using System.Data;
using Aspose.Words;

class MailMergeToXps
{
    static void Main()
    {
        // Path to the source DOCX file that contains MERGEFIELDs.
        string docPath = @"C:\Input\Template.docx";

        // Path to the XML file that holds the mail‑merge data.
        string xmlPath = @"C:\Input\Data.xml";

        // Load the DOCX document.
        Document doc = new Document(docPath);

        // Load the XML data into a DataSet.  Aspose.Words can merge from a DataSet,
        // and DataSet can read XML directly.
        DataSet dataSet = new DataSet();
        dataSet.ReadXml(xmlPath);

        // Perform the mail merge using the first table in the DataSet.
        // If the XML contains multiple tables you can choose the appropriate one.
        doc.MailMerge.Execute(dataSet.Tables[0]);

        // Save the merged document as XPS.
        string outPath = @"C:\Output\MergedDocument.xps";
        doc.Save(outPath, SaveFormat.Xps);
    }
}
