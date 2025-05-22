using ClosedXML.Excel;
using Hotcakes.CommerceDTO.v1;
using Hotcakes.CommerceDTO.v1.Catalog;
using Hotcakes.CommerceDTO.v1.Client;
using Hotcakes.CommerceDTO.v1.Contacts;
using Hotcakes.CommerceDTO.v1.Orders;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO.Packaging;
using System.Linq;
using System.Windows.Forms;
using System.Xml;



namespace Rendeleskezelo
{


    public partial class MainForm : Form
    {

        List<string> products = new List<string>();
        List<int> quantities = new List<int>();
        List<decimal> prices = new List<decimal>();
        string orderId = string.Empty;
        string xml = string.Empty;
        public MainForm()
        {
            InitializeComponent();
            this.StartPosition = FormStartPosition.CenterScreen;
            dataGridViewOrders.DataSource = orderDTOBindingSource;
            BetoltesRendelesek();
            //PopulateStatusComboBoxFromApi();
            orderId = ((OrderDTO)orderDTOBindingSource.Current).bvin;
        }
        public DataGridView OrdersGrid
        {
            get { return dataGridViewOrders; }
        }


        private static Api ApiHivas()
        {
            string url = "http://rendfejl1016.northeurope.cloudapp.azure.com:8080";
            string kulcs = "1-9b1353c9-35ee-49e6-bfd2-cde51b368def";
            Api proxy = new Api(url, kulcs);
            return proxy;
        }

        private void BetoltesRendelesek()
        {
            Api proxy = ApiHivas();
            var response = proxy.OrdersFindAll();
            if (response == null || response.Content == null || response.Content.Count == 0)
            {
                MessageBox.Show("Nem sikerült lekérni a rendeléseket.");
                return;
            }

            string filterCustomerName = textBoxCustomerName.Text.Trim().ToLower();

            BindingList<OrderDTO> filteredOrders = new BindingList<OrderDTO>(
                response.Content
                    .Where(order => !string.IsNullOrEmpty(order.StatusName) &&
                                    (string.IsNullOrEmpty(filterCustomerName) ||
                                     (order.BillingAddress.FirstName + " " + order.BillingAddress.LastName).ToLower().Contains(filterCustomerName))
                                    )
                    .Select(order => new OrderDTO(order))  // Map to custom DTO
                    .ToList()
            );

            orderDTOBindingSource.DataSource = filteredOrders;
            LoadOrders(proxy);

        }





        private void textBoxCustomerName_TextChanged(object sender, EventArgs e)
        {
            BetoltesRendelesek();
        }




        private void buttonReload_Click(object sender, EventArgs e)
        {
            BetoltesRendelesek();
        }     

        private void MainForm_Load(object sender, EventArgs e)
        {
        }

        private void dataGridViewOrders_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            string columnName = dataGridViewOrders.Columns[e.ColumnIndex].DataPropertyName;
            var direction = ListSortDirection.Ascending;

            // Optional: toggle direction if already sorted
            if (orderDTOBindingSource.Sort == $"{columnName} ASC")
                direction = ListSortDirection.Descending;

            orderDTOBindingSource.Sort = $"{columnName} {(direction == ListSortDirection.Ascending ? "ASC" : "DESC")}";
        }

        private void buttonMinta_Click(object sender, EventArgs e)
        {
            GenerateXML();
        }
        //public void GenerateXMLToFile(string filePath)
        //{
        //    using (XmlWriter writer = XmlWriter.Create(filePath))
        //    {
        //        writer.WriteStartDocument();
        //        writer.WriteStartElement("Orders"); // Root element

        //        foreach (DataGridViewRow row in dataGridViewOrders.Rows)
        //        {
        //            if (row.IsNewRow) continue;

        //            var cell = row.Cells["checkBox"].Value;
        //            bool isChecked = cell != null && Convert.ToBoolean(cell);

        //            if (isChecked)
        //            {
        //                string orderID = row.Cells["orderIDDataGridViewTextBoxColumn"].Value?.ToString() ?? "";
        //                string customerName = row.Cells["customerNameDataGridViewTextBoxColumn"].Value?.ToString() ?? "";
        //                string orderDate = row.Cells["orderDateDataGridViewTextBoxColumn"].Value?.ToString() ?? "";
        //                string address = row.Cells["shippingAddressDataGridViewTextBoxColumn"].Value?.ToString() ?? "";
        //                string orderTotal = row.Cells["orderTotalDataGridViewTextBoxColumn"].Value?.ToString() ?? "";

        //                writer.WriteStartElement("Order");

        //                writer.WriteElementString("customerName", customerName);
        //                writer.WriteElementString("customerAddress", address);
        //                writer.WriteElementString("invoiceDeliveryDate", orderDate);
        //                writer.WriteElementString("productCodeValue", orderID);
        //                writer.WriteElementString("unitPrice", orderTotal);
        //                writer.WriteElementString("unitPriceHUF", orderTotal);

        //                writer.WriteEndElement(); // </Order>
        //            }
        //        }

        //        writer.WriteEndElement(); // </Orders>
        //        writer.WriteEndDocument();
        //    }
        //}

        public void GenerateXML()
        {
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {

                    // Save the file to the selected location
                    using (XmlWriter writer = XmlWriter.Create("Szamla.xml"))
                    {
                        writer.WriteStartDocument();
                        writer.WriteComment("");
                        writer.WriteStartElement("Orders");
                        writer.WriteStartElement("invoiceMain");
                        writer.WriteStartElement("invoice");
                        writer.WriteStartElement("invoiceHead");
                        writer.WriteStartElement("supplierInfo");
                        writer.WriteStartElement("supplierTaxNumber");
                        writer.WriteElementString("taxpayerId", "99999999");
                        writer.WriteElementString("vatCode", "2");
                        writer.WriteElementString("countyCode", "41");

                        foreach (DataGridViewRow row in dataGridViewOrders.Rows)
                        {
                            if (row.IsNewRow) continue;

                            var cell = row.Cells["checkBox"].Value;
                            bool isChecked = cell != null && Convert.ToBoolean(cell);

                            if (isChecked)
                            {
                                string orderID = row.Cells["orderIDDataGridViewTextBoxColumn"].Value?.ToString() ?? "";
                                string customerName = row.Cells["customerNameDataGridViewTextBoxColumn"].Value?.ToString() ?? "";
                                string orderDate = row.Cells["orderDateDataGridViewTextBoxColumn"].Value?.ToString() ?? "";
                                string address = row.Cells["shippingAddressDataGridViewTextBoxColumn"].Value?.ToString() ?? "";
                                string orderTotal = row.Cells["orderTotalDataGridViewTextBoxColumn"].Value?.ToString() ?? "";

                                //MessageBox.Show($"kijelölt rendelés: {orderID}; {customerName}");

                                // Write XML entry for each selected row

                                writer.WriteComment("");
                                writer.WriteStartElement("Order");

                                writer.WriteElementString("customerName", customerName);
                                writer.WriteElementString("customerAddress", address);
                                writer.WriteElementString("invoiceDeliveryDate", orderDate);
                                writer.WriteElementString("productCodeValue", orderID);
                                writer.WriteElementString("unitPrice", orderTotal);
                                writer.WriteElementString("unitPriceHUF", orderTotal);

                                writer.WriteEndElement(); // </Order>
                            }
                        }
                        
                        writer.WriteEndElement(); // </Orders>
                        writer.WriteEndDocument();
                       
                    }

                    MessageBox.Show("XML sikeresen mentve: szamla.xml");
                }
            }

        }

  

        private void LoadOrders(Api proxy)
        {
            orderId = ((OrderDTO)orderDTOBindingSource.Current).bvin;
            var order = proxy.OrdersFind(orderId);

            products = order.Content.Items.Select(item => item.ProductName).ToList();
            quantities = order.Content.Items.Select(item => item.Quantity).ToList();
            prices = order.Content.Items.Select(item => item.BasePricePerItem).ToList();
        }

        private void dataGridViewOrders_SelectionChanged(object sender, EventArgs e)
        {
            Api proxy = ApiHivas();
            LoadOrders(proxy);
        }


        private void dataGridViewOrders_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {

        }

    }
}
