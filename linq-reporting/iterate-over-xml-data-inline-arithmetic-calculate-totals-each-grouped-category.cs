using System;
using System.Linq;
using System.Xml.Linq;

namespace LinqReporting
{
    class Program
    {
        static void Main()
        {
            // Sample XML data.
            string xml = @"<Report>
  <Category name='Food'>
    <Item amount='10.5' tax='0.5' />
    <Item amount='20.0' tax='1.0' />
  </Category>
  <Category name='Books'>
    <Item amount='15.0' tax='0.75' />
  </Category>
</Report>";

            // Load the XML into an XDocument.
            XDocument doc = XDocument.Parse(xml);

            // Group items by category and calculate totals using inline arithmetic.
            var categoryTotals = doc.Root
                .Elements("Category")
                .Select(cat => new
                {
                    Name = (string)cat.Attribute("name"),
                    Total = cat.Elements("Item")
                               .Sum(item => (double)item.Attribute("amount") + (double)item.Attribute("tax"))
                });

            // Output the results.
            foreach (var ct in categoryTotals)
            {
                Console.WriteLine($"{ct.Name}: {ct.Total}");
            }
        }
    }
}
