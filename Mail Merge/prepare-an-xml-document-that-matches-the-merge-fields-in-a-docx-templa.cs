using System;
using System.Xml;
using Aspose.Words;

namespace MailMergeXmlGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the DOCX template that contains MERGEFIELDs.
            const string templatePath = @"C:\Docs\Template.docx";

            // Load the template document.
            Document template = new Document(templatePath);

            // Retrieve all merge field names from the template.
            string[] fieldNames = template.MailMerge.GetFieldNames();

            // Create an XML document that will hold the data for the merge.
            XmlDocument xmlData = new XmlDocument();

            // Create the XML declaration.
            XmlDeclaration xmlDecl = xmlData.CreateXmlDeclaration("1.0", "UTF-8", null);
            xmlData.AppendChild(xmlDecl);

            // Create the root element (e.g., <Data>).
            XmlElement root = xmlData.CreateElement("Data");
            xmlData.AppendChild(root);

            // For each merge field, add an element with a placeholder value.
            foreach (string fieldName in fieldNames)
            {
                // Create an element named after the field.
                XmlElement fieldElement = xmlData.CreateElement(fieldName);

                // Set a sample placeholder value; adjust as needed.
                fieldElement.InnerText = $"Sample value for {fieldName}";

                // Append the field element to the root.
                root.AppendChild(fieldElement);
            }

            // Path where the generated XML will be saved.
            const string xmlOutputPath = @"C:\Docs\MergeData.xml";

            // Save the XML document.
            xmlData.Save(xmlOutputPath);

            Console.WriteLine($"XML data file created at: {xmlOutputPath}");
        }
    }
}
