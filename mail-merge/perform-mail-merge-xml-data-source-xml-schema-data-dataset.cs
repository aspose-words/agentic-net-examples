using System;
using System.Data;
using Aspose.Words;

class MailMergeXmlExample
{
    static void Main()
    {
        // Paths to the template, XML data, XML schema and the output document.
        string templatePath = "Template.docx";
        string xmlDataPath = "Data.xml";
        string xmlSchemaPath = "Schema.xsd";
        string outputPath = "MergedOutput.docx";

        // Load the Word template that contains mail‑merge regions.
        Document doc = new Document(templatePath);

        // Create a DataSet and populate it with the XML schema and data.
        DataSet dataSet = new DataSet();

        // Load the XSD schema first so the DataSet knows the table structure.
        dataSet.ReadXmlSchema(xmlSchemaPath);

        // Load the XML data into the DataSet according to the previously loaded schema.
        dataSet.ReadXml(xmlDataPath);

        // Execute mail merge using the DataSet. The DataSet tables must have names that match
        // the mail‑merge region names defined in the template (TableStart:TableName / TableEnd:TableName).
        doc.MailMerge.ExecuteWithRegions(dataSet);

        // Save the merged document.
        doc.Save(outputPath);
    }
}
