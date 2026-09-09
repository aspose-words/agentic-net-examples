using System;
using System.Data;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare a simple XML file that represents a DataSet with one table named "Customers".
        string xmlContent =
            @"<?xml version=""1.0"" encoding=""utf-8""?>"
          + @"<DataSet>"
          + @"  <Customers>"
          + @"    <Customer>"
          + @"      <CustomerName>Thomas Hardy</CustomerName>"
          + @"      <Address>120 Hanover Sq., London</Address>"
          + @"    </Customer>"
          + @"    <Customer>"
          + @"      <CustomerName>Paolo Accorti</CustomerName>"
          + @"      <Address>Via Monte Bianco 34, Torino</Address>"
          + @"    </Customer>"
          + @"  </Customers>"
          + @"</DataSet>";

        // Write the XML to a temporary file.
        string xmlPath = Path.Combine(Path.GetTempPath(), "Customers.xml");
        File.WriteAllText(xmlPath, xmlContent);

        // Load the XML data into a DataSet using ReadXml.
        DataSet dataSet = new DataSet();
        dataSet.ReadXml(xmlPath);

        // Create a new blank Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a mail‑merge region that corresponds to the "Customers" table.
        // TableStart and TableEnd fields define the region boundaries.
        builder.InsertField(" MERGEFIELD TableStart:Customers");
        builder.InsertField(" MERGEFIELD CustomerName");
        builder.Write(" - ");
        builder.InsertField(" MERGEFIELD Address");
        builder.InsertField(" MERGEFIELD TableEnd:Customers");

        // Perform mail merge using the DataSet that was loaded from XML.
        // ExecuteWithRegions will process the region defined above.
        doc.MailMerge.ExecuteWithRegions(dataSet);

        // Save the resulting document.
        string outputPath = Path.Combine(Path.GetTempPath(), "MailMergeResult.docx");
        doc.Save(outputPath);

        // Inform the user where the file was saved.
        Console.WriteLine("Mail merge completed. Document saved to:");
        Console.WriteLine(outputPath);
    }
}
