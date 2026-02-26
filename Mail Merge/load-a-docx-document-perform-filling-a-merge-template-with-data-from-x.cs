using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the Word template that contains MERGEFIELD tags (and optional mail‑merge regions).
        Document doc = new Document("Template.docx");

        // Load XML data into a DataSet. The XML file should have a structure that matches the
        // merge fields or regions defined in the template.
        DataSet dataSet = new DataSet();
        dataSet.ReadXml("Data.xml");

        // Fill the template with the data from the DataSet.
        // ExecuteWithRegions is used so that, if the template contains mail‑merge regions,
        // they are processed correctly; otherwise Execute would also work.
        doc.MailMerge.ExecuteWithRegions(dataSet);

        // Save the resulting document as XPS.
        XpsSaveOptions xpsOptions = new XpsSaveOptions(); // default XPS options
        doc.Save("Result.xps", xpsOptions);
    }
}
