using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using Aspose.Words;
using Aspose.Words.MailMerging;
using Aspose.Words.Reporting;

namespace MailMergeWithLinqReporting
{
    // Simple data entity.
    public class Customer
    {
        public string FullName { get; set; }
        public string Address { get; set; }
        public string City { get; set; }

        public Customer(string fullName, string address, string city)
        {
            FullName = fullName;
            Address = address;
            City = city;
        }
    }

    // Custom mail‑merge data source that wraps a list of Customer objects.
    public class CustomerMailMergeDataSource : IMailMergeDataSource
    {
        private readonly IList<Customer> _customers;
        private int _recordIndex = -1;

        public CustomerMailMergeDataSource(IList<Customer> customers)
        {
            _customers = customers;
        }

        // The name of the data source – must match the mail‑merge region name if regions are used.
        public string TableName => "Customer";

        // Returns the value for a given field name.
        public bool GetValue(string fieldName, out object fieldValue)
        {
            switch (fieldName)
            {
                case "FullName":
                    fieldValue = _customers[_recordIndex].FullName;
                    return true;
                case "Address":
                    fieldValue = _customers[_recordIndex].Address;
                    return true;
                case "City":
                    fieldValue = _customers[_recordIndex].City;
                    return true;
                default:
                    fieldValue = null;
                    return false;
            }
        }

        // Moves to the next record; returns false when the end of the collection is reached.
        public bool MoveNext()
        {
            if (_recordIndex < _customers.Count - 1)
            {
                _recordIndex++;
                return true;
            }
            return false;
        }

        // No child data sources are required for this simple example.
        public IMailMergeDataSource GetChildDataSource(string tableName) => null;
    }

    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Prepare sample data – a list of customers.
            // -----------------------------------------------------------------
            List<Customer> customers = new List<Customer>
            {
                new Customer("Thomas Hardy", "120 Hanover Sq.", "London"),
                new Customer("Paolo Accorti", "Via Monte Bianco 34", "Torino")
            };

            // -----------------------------------------------------------------
            // 2. Use the ReportingEngine (LINQ based) to fill a template that
            //    contains ReportingEngine syntax (<<[customer.FullName]>> etc.).
            // -----------------------------------------------------------------
            Document reportingTemplate = CreateReportingTemplate();
            ReportingEngine reportingEngine = new ReportingEngine();

            // The ReportingEngine can work directly with any enumerable object.
            // Here we pass the list of customers; the engine will iterate over it.
            reportingEngine.BuildReport(reportingTemplate, customers);
            reportingTemplate.Save("ReportingResult.docx");

            // -----------------------------------------------------------------
            // 3. Use the classic MailMerge engine with the same data source.
            //    The template contains MERGEFIELDs that correspond to the field
            //    names defined in CustomerMailMergeDataSource.
            // -----------------------------------------------------------------
            Document mailMergeTemplate = CreateMailMergeTemplate();
            CustomerMailMergeDataSource dataSource = new CustomerMailMergeDataSource(customers);
            mailMergeTemplate.MailMerge.Execute(dataSource);
            mailMergeTemplate.Save("MailMergeResult.docx");
        }

        // Creates a template for the ReportingEngine.
        // The template uses the <<[customer.Property]>> syntax.
        private static Document CreateReportingTemplate()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Customer Report");
            builder.Writeln("----------------");
            // The ReportingEngine will repeat the block for each item in the collection.
            builder.Writeln("<<foreach [customer]>>");
            builder.Writeln("Name   : <<[customer.FullName]>>");
            builder.Writeln("Address: <<[customer.Address]>>");
            builder.Writeln("City   : <<[customer.City]>>");
            builder.Writeln("<<end>>");

            return doc;
        }

        // Creates a template for the classic MailMerge engine.
        // MERGEFIELD names must match the values returned by GetValue().
        private static Document CreateMailMergeTemplate()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Customer List");
            builder.Writeln("==============");
            builder.InsertField(" MERGEFIELD FullName ");
            builder.Write(", ");
            builder.InsertField(" MERGEFIELD Address ");
            builder.Write(", ");
            builder.InsertField(" MERGEFIELD City ");
            builder.Writeln();

            return doc;
        }
    }
}
