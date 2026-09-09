using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class MailMergeFromXml
{
    public static void Main()
    {
        // Define file paths in the current directory.
        string xmlPath = Path.Combine(Environment.CurrentDirectory, "Data.xml");
        string templatePath = Path.Combine(Environment.CurrentDirectory, "Template.docx");
        string outputPath = Path.Combine(Environment.CurrentDirectory, "MergedOutput.docx");

        // Create a simple XML data source.
        string xmlContent = @"
<Root>
    <Person>
        <FirstName>John</FirstName>
        <LastName>Doe</LastName>
        <Message>Hello! This is a merged message.</Message>
    </Person>
    <Person>
        <FirstName>Jane</FirstName>
        <LastName>Smith</LastName>
        <Message>Welcome to Aspose.Words mail merge.</Message>
    </Person>
</Root>";
        File.WriteAllText(xmlPath, xmlContent);

        // Build a mail‑merge template document with a region named "Person".
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Begin the region.
        builder.InsertField(" MERGEFIELD TableStart:Person");
        // Insert the fields that will be filled from the XML.
        builder.Write("First Name: ");
        builder.InsertField(" MERGEFIELD FirstName");
        builder.Writeln();
        builder.Write("Last Name: ");
        builder.InsertField(" MERGEFIELD LastName");
        builder.Writeln();
        builder.Write("Message: ");
        builder.InsertField(" MERGEFIELD Message");
        builder.Writeln();
        // End the region.
        builder.InsertField(" MERGEFIELD TableEnd:Person");

        // Save the template to disk (required by the rule to use a save operation).
        template.Save(templatePath);

        // Load the XML into a DataSet.
        DataSet dataSet = new DataSet();
        dataSet.ReadXml(xmlPath);

        // Load the template document (required by the rule to use a load operation).
        Document doc = new Document(templatePath);

        // Perform mail merge using the DataSet. The DataSet contains a table named "Person"
        // which matches the region name in the template.
        doc.MailMerge.ExecuteWithRegions(dataSet);

        // Save the merged document.
        doc.Save(outputPath);
    }
}
