using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Order
{
    public string CustomerName { get; set; } = "";
    public DateTime OrderDate { get; set; }

    public Order(string customerName, DateTime orderDate)
    {
        CustomerName = customerName;
        OrderDate = orderDate;
    }
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        List<Order> allOrders = new List<Order>
        {
            new Order("Alice", DateTime.Today.AddDays(-5)),          // within current month
            new Order("Bob",   DateTime.Today.AddMonths(-1).AddDays(-2)), // within last month
            new Order("Carol", DateTime.Today.AddMonths(-1).AddDays(-15)),// within last month
            new Order("Dave",  DateTime.Today.AddMonths(-2)),       // older than last month
            new Order("Eve",   DateTime.Today)                     // today
        };

        // Determine the date range for the previous month.
        DateTime today = DateTime.Today;
        DateTime firstDayOfCurrentMonth = new DateTime(today.Year, today.Month, 1);
        DateTime startOfLastMonth = firstDayOfCurrentMonth.AddMonths(-1);
        DateTime endOfLastMonth = firstDayOfCurrentMonth;

        // Filter orders using a lambda expression inside Where.
        List<Order> filteredOrders = allOrders
            .Where(o => o.OrderDate >= startOfLastMonth && o.OrderDate < endOfLastMonth)
            .ToList();

        // Create a template document programmatically.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Orders from the previous month:");
        builder.Writeln("<<foreach [order in orders]>>");
        builder.Writeln("- <<[order.CustomerName]>> placed on <<[order.OrderDate]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk as required by the workflow.
        string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // Load the template back before building the report.
        Document reportDoc = new Document(templatePath);

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        // The root object name must match the tag reference ("orders").
        engine.BuildReport(reportDoc, filteredOrders, "orders");

        // Save the final report.
        string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
