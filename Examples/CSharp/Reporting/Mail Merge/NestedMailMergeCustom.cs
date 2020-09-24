using System.Collections;
using Aspose.Words.MailMerging;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp
{
    class NestedMailMergeCustom : TestDataHelper
    {
        [Test]
        public static void Run()
        {
            //ExStart:NestedMailMergeCustom
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertField(" MERGEFIELD TableStart:Customer");

            builder.Write("Full name:\t");
            builder.InsertField(" MERGEFIELD FullName ");
            builder.Write("\nAddress:\t");
            builder.InsertField(" MERGEFIELD Address ");
            builder.Write("\nOrders:\n");

            builder.InsertField(" MERGEFIELD TableStart:Order");

            builder.Write("\tItem name:\t");
            builder.InsertField(" MERGEFIELD Name ");
            builder.Write("\n\tQuantity:\t");
            builder.InsertField(" MERGEFIELD Quantity ");
            builder.InsertParagraph();

            builder.InsertField(" MERGEFIELD TableEnd:Order");

            builder.InsertField(" MERGEFIELD TableEnd:Customer");

            // Create some data that we will use in the mail merge
            CustomerList customers = new CustomerList();
            customers.Add(new Customer("Thomas Hardy", "120 Hanover Sq., London"));
            customers.Add(new Customer("Paolo Accorti", "Via Monte Bianco 34, Torino"));

            // Create some data for nesting in the mail merge
            customers[0].Orders.Add(new Order("Rugby World Cup Cap", 2));
            customers[0].Orders.Add(new Order("Rugby World Cup Ball", 1));
            customers[1].Orders.Add(new Order("Rugby World Cup Guide", 1));

            // To be able to mail merge from your own data source, it must be wrapped
            // Into an object that implements the IMailMergeDataSource interface
            CustomerMailMergeDataSource customersDataSource = new CustomerMailMergeDataSource(customers);

            // Now you can pass your data source into Aspose.Words
            doc.MailMerge.ExecuteWithRegions(customersDataSource);

            doc.Save(ArtifactsDir + "MailMerge.NestedMailMergeCustom.docx");
            //ExEnd:NestedMailMergeCustom
        }

        /// <summary>
        /// An example of a "data entity" class in your application.
        /// </summary>
        public class Customer
        {
            public Customer(string aFullName, string anAddress)
            {
                FullName = aFullName;
                Address = anAddress;
                Orders = new OrderList();
            }

            public string FullName { get; set; }
            public string Address { get; set; }
            public OrderList Orders { get; set; }
        }

        /// <summary>
        /// An example of a typed collection that contains your "data" objects.
        /// </summary>
        public class CustomerList : ArrayList
        {
            public new Customer this[int index]
            {
                get => (Customer) base[index];
                set => base[index] = value;
            }
        }

        /// <summary>
        /// An example of a child "data entity" class in your application.
        /// </summary>
        public class Order
        {
            public Order(string oName, int oQuantity)
            {
                Name = oName;
                Quantity = oQuantity;
            }

            public string Name { get; set; }
            public int Quantity { get; set; }
        }

        /// <summary>
        /// An example of a typed collection that contains your "data" objects.
        /// </summary>
        public class OrderList : ArrayList
        {
            public new Order this[int index]
            {
                get => (Order) base[index];
                set => base[index] = value;
            }
        }

        /// <summary>
        /// A custom mail merge data source that you implement to allow Aspose.Words 
        /// To mail merge data from your Customer objects into Microsoft Word documents.
        /// </summary>
        public class CustomerMailMergeDataSource : IMailMergeDataSource
        {
            public CustomerMailMergeDataSource(CustomerList customers)
            {
                mCustomers = customers;

                // When the data source is initialized, it must be positioned before the first record
                mRecordIndex = -1;
            }

            /// <summary>
            /// The name of the data source. Used by Aspose.Words only when executing mail merge with repeatable regions.
            /// </summary>
            public string TableName => "Customer";

            /// <summary>
            /// Aspose.Words calls this method to get a value for every data field.
            /// </summary>
            public bool GetValue(string fieldName, out object fieldValue)
            {
                switch (fieldName)
                {
                    case "FullName":
                        fieldValue = mCustomers[mRecordIndex].FullName;
                        return true;
                    case "Address":
                        fieldValue = mCustomers[mRecordIndex].Address;
                        return true;
                    case "Order":
                        fieldValue = mCustomers[mRecordIndex].Orders;
                        return true;
                    default:
                        // A field with this name was not found,
                        // Return false to the Aspose.Words mail merge engine
                        fieldValue = null;
                        return false;
                }
            }

            /// <summary>
            /// A standard implementation for moving to a next record in a collection.
            /// </summary>
            public bool MoveNext()
            {
                if (!IsEof)
                    mRecordIndex++;

                return !IsEof;
            }

            //ExStart:GetChildDataSourceExample           
            public IMailMergeDataSource GetChildDataSource(string tableName)
            {
                switch (tableName)
                {
                    // Get the child collection to merge it with the region provided with tableName variable.
                    case "Order":
                        return new OrderMailMergeDataSource(mCustomers[mRecordIndex].Orders);
                    default:
                        return null;
                }
            }
            //ExEnd:GetChildDataSourceExample

            private bool IsEof => (mRecordIndex >= mCustomers.Count);

            private readonly CustomerList mCustomers;
            private int mRecordIndex;
        }

        public class OrderMailMergeDataSource : IMailMergeDataSource
        {
            public OrderMailMergeDataSource(OrderList orders)
            {
                mOrders = orders;

                // When the data source is initialized, it must be positioned before the first record
                mRecordIndex = -1;
            }

            /// <summary>
            /// The name of the data source. Used by Aspose.Words only when executing mail merge with repeatable regions.
            /// </summary>
            public string TableName => "Order";

            /// <summary>
            /// Aspose.Words calls this method to get a value for every data field.
            /// </summary>
            public bool GetValue(string fieldName, out object fieldValue)
            {
                switch (fieldName)
                {
                    case "Name":
                        fieldValue = mOrders[mRecordIndex].Name;
                        return true;
                    case "Quantity":
                        fieldValue = mOrders[mRecordIndex].Quantity;
                        return true;
                    default:
                        // A field with this name was not found,
                        // Return false to the Aspose.Words mail merge engine
                        fieldValue = null;
                        return false;
                }
            }

            /// <summary>
            /// A standard implementation for moving to a next record in a collection.
            /// </summary>
            public bool MoveNext()
            {
                if (!IsEof)
                    mRecordIndex++;

                return !IsEof;
            }

            // Return null because we haven't any child elements for this sort of object
            public IMailMergeDataSource GetChildDataSource(string tableName)
            {
                return null;
            }

            private bool IsEof => mRecordIndex >= mOrders.Count;

            private readonly OrderList mOrders;
            private int mRecordIndex;
        }
    }
}