using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Path for temporary files.
        string xmlPath = "people.xml";
        string outputPath = "MailMergeResult.docx";

        // Create a simple XML file that represents a list of persons.
        // The root element contains multiple <person> elements.
        string xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<persons>
    <person>
        <FullName>Thomas Hardy</FullName>
        <Address>120 Hanover Sq., London</Address>
    </person>
    <person>
        <FullName>Paolo Accorti</FullName>
        <Address>Via Monte Bianco 34, Torino</Address>
    </person>
</persons>";
        File.WriteAllText(xmlPath, xmlContent);

        // Load the XML data into a DataSet using ReadXml.
        DataSet dataSet = new DataSet();
        dataSet.ReadXml(xmlPath);

        // Build a mail‑merge template document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define a mail‑merge region named "persons".
        builder.InsertField(" MERGEFIELD TableStart:persons");
        builder.InsertField(" MERGEFIELD FullName");
        builder.Write("\t");
        builder.InsertField(" MERGEFIELD Address");
        builder.InsertParagraph();
        builder.InsertField(" MERGEFIELD TableEnd:persons");

        // Perform mail merge using the DataSet.
        // The DataSet contains a table named "person" that matches the region name (persons) by default.
        // Aspose.Words will map the table "person" to the region "persons".
        doc.MailMerge.ExecuteWithRegions(dataSet);

        // Save the merged document.
        doc.Save(outputPath);

        // Clean up temporary XML file.
        File.Delete(xmlPath);
    }
}
