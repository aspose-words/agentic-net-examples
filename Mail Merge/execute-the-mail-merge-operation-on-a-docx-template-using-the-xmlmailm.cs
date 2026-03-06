using System;
using System.Collections;
using System.Collections.Generic;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.MailMerging;

namespace MailMergeExample
{
    // Custom data source that reads merge data from an XML file.
    // Implements IMailMergeDataSource so Aspose.Words can perform the merge.
    public class XmlMailMergeDataSource : IMailMergeDataSource
    {
        private readonly List<XElement> _records;   // Each element represents one record.
        private int _recordIndex = -1;               // Position before the first record.

        // The name of the data source (used for mail‑merge regions).
        public string TableName => "Customer";

        // Constructor loads the XML file and extracts the record elements.
        public XmlMailMergeDataSource(string xmlFilePath)
        {
            // Expected XML format:
            // <Customers>
            //   <Customer>
            //     <FullName>John Doe</FullName>
            //     <Address>123 Main St.</Address>
            //   </Customer>
            //   ...
            // </Customers>
            XDocument doc = XDocument.Load(xmlFilePath);
            _records = new List<XElement>(doc.Root.Elements("Customer"));
        }

        // Moves to the next record. Returns false when no more records are available.
        public bool MoveNext()
        {
            if (_recordIndex + 1 < _records.Count)
            {
                _recordIndex++;
                return true;
            }
            return false;
        }

        // Retrieves the value for a given field name from the current record.
        public bool GetValue(string fieldName, out object fieldValue)
        {
            // Look for a child element with the same name as the field.
            XElement current = _records[_recordIndex];
            XElement element = current.Element(fieldName);
            if (element != null)
            {
                fieldValue = element.Value;
                return true;
            }

            fieldValue = null;
            return false; // Field not found.
        }

        // No child data sources are required for this simple example.
        public IMailMergeDataSource GetChildDataSource(string tableName) => null;
    }

    class Program
    {
        static void Main()
        {
            // Path to the DOCX template that contains MERGEFIELDs: FullName and Address.
            const string templatePath = "Template.docx";

            // Path to the XML file that holds the data.
            const string xmlDataPath = "Customers.xml";

            // Load the template document.
            Document doc = new Document(templatePath);

            // Create the custom XML data source.
            IMailMergeDataSource dataSource = new XmlMailMergeDataSource(xmlDataPath);

            // Execute the mail merge using the custom data source.
            doc.MailMerge.Execute(dataSource);

            // Save the merged document.
            const string outputPath = "MergedResult.docx";
            doc.Save(outputPath);
        }
    }
}
