using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a simple template document with merge fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Dear <<FirstName>> <<LastName>>,");
        builder.Writeln("Your order <<OrderID>> is confirmed.");
        builder.Writeln("Thank you for shopping with us.");

        // XML data to be used as the mail merge source.
        string xmlData = @"
<Customers>
    <Customer>
        <FirstName>John</FirstName>
        <LastName>Doe</LastName>
        <OrderID>12345</OrderID>
    </Customer>
    <Customer>
        <FirstName>Jane</FirstName>
        <LastName>Smith</LastName>
        <OrderID>67890</OrderID>
    </Customer>
</Customers>";

        // Load XML into a DataSet using ReadXml.
        DataSet dataSet = new DataSet();
        using (StringReader sr = new StringReader(xmlData))
        {
            dataSet.ReadXml(sr);
        }

        // Perform mail merge using the first table in the DataSet.
        if (dataSet.Tables.Count > 0)
        {
            doc.MailMerge.Execute(dataSet.Tables[0]);
        }

        // Save the merged document.
        string outputPath = "MergedOutput.docx";
        doc.Save(outputPath);

        // Indicate completion.
        Console.WriteLine($"Mail merge completed. Document saved to '{Path.GetFullPath(outputPath)}'.");
    }
}
