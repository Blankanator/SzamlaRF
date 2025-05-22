using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Vml.Office;
using Hotcakes.CommerceDTO.v1;
using Hotcakes.CommerceDTO.v1.Catalog;
using Hotcakes.CommerceDTO.v1.Client;
using Hotcakes.CommerceDTO.v1.Orders;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Rendeleskezeles
{
    public class ProductDTO
    {
        public string ProductName { get; set; }
    }
    public partial class ProductsForm : Form
    {
        public static List<ProductDTO> FilterProducts(List<ProductDTO> allProducts, string filterText)
        {
            return allProducts
                .Where(p => p.ProductName.ToLower().Contains(filterText.ToLower()))
                .ToList();
        }

        List<string> items = new List<string>();
        List<int> quantity = new List<int>();
        string orderId = string.Empty;
        public ProductsForm(BindingSource bindingSource)
        {
            InitializeComponent();
            this.StartPosition = FormStartPosition.CenterScreen;
            this.orderDTOBindingSource.DataSource = bindingSource.Current;
            orderId = ((OrderDTO)bindingSource.Current).bvin;
        }

        private static Api ApiHivas()
        {
            string url = "http://rendfejl1016.northeurope.cloudapp.azure.com:8080";
            string kulcs = "1-9b1353c9-35ee-49e6-bfd2-cde51b368def";
            Api proxy = new Api(url, kulcs);
            return proxy;
        }

        private void ProductsForm_Load(object sender, EventArgs e)
        {
            foreach (System.Windows.Forms.Control ctrl in this.Controls)
            {
                if (ctrl is Button btn)
                {
                    btn.FlatStyle = FlatStyle.Flat;
                    btn.BackColor = System.Drawing.Color.DeepSkyBlue;
                    btn.FlatAppearance.BorderColor = System.Drawing.Color.DeepSkyBlue;
                    btn.FlatAppearance.BorderSize = 0;
                    btn.FlatAppearance.MouseOverBackColor = System.Drawing.Color.DeepSkyBlue;
                    btn.FlatAppearance.MouseDownBackColor = System.Drawing.Color.DeepSkyBlue;
                }
            }

            Api proxy = ApiHivas();

            var response = proxy.ProductsFindAll();

            LoadOrders(proxy);

            productBindingSource.DataSource = response.Content.ToList();
            listBoxProducts.DataSource = productBindingSource;
            listBoxProducts.DisplayMember = "ProductName";
        }

        private void textBoxFilter_TextChanged(object sender, EventArgs e)
        {
            string filterText = textBoxFilter.Text.ToLower();

            var allProducts = (List<ProductDTO>)productBindingSource.DataSource;

            var filtered = allProducts
            .Where(p => p.ProductName.ToLower().Contains(filterText))
            .ToList();

            listBoxProducts.DataSource = null;
            listBoxProducts.DataSource = filtered;
            listBoxProducts.DisplayMember = "ProductName";
        }

        private void buttonRemove_Click(object sender, EventArgs e)
        {
            UpdateOrder_Remove();
        }

        private void buttonAdd_Click(object sender, EventArgs e)
        {
            UpdateOrder_Add();
        }

        private void listBoxOrdered_SelectedIndexChanged(object sender, EventArgs e)
        {
            int selectedIndex = listBoxOrdered.SelectedIndex;

            if (selectedIndex >= 0 && selectedIndex < quantity.Count)
            {
                labelQuantity.Text = $"Mennyiség: {quantity[selectedIndex]}";
            }
            else
            {
                labelQuantity.Text = "";
            }
        }

        private void UpdateOrder_Add()
        {
            Api proxy = ApiHivas();
            var response = proxy.OrdersFind(orderId);
            var originalOrder = response.Content;
            if (numericQuantity.Value <= 0 || numericQuantity.Value % 1 != 0)
            {
                MessageBox.Show("Kérjük, adjon meg egy érvényes mennyiséget a termékhez.");
                return;
            }
            int quantity = (int)numericQuantity.Value;
            string selectedProductName = listBoxProducts.Text?.ToString();
            if (string.IsNullOrEmpty(selectedProductName))
            {
                MessageBox.Show("Nincs kiválasztva termék a rendelésből.");
                return;
            }

            // Betöltjük az összes terméket és megkeressük a nevet
            var allProducts = proxy.ProductsFindAll();
            var matchingProduct = allProducts.Content.FirstOrDefault(p => p.ProductName == selectedProductName);

            if (matchingProduct == null)
            {
                MessageBox.Show($"Nem található termék ilyen névvel: {selectedProductName}");
                return;
            }

            string productId = matchingProduct.Bvin;

            var product = proxy.ProductsFind(productId);

            if (product.Content == null)
            {
                MessageBox.Show("Nem található termék azonosítóval: " + productId);
                return;
            }

            originalOrder.Items.Add(new LineItemDTO
            {
                ProductId = productId,
                Quantity = quantity,
                ProductShortDescription = product.Content.ShortDescription,
                ProductName = product.Content.ProductName,
                ProductSku = product.Content.Sku,
                BasePricePerItem = product.Content.SitePrice,
                LineTotal = product.Content.SitePrice * quantity
            });

            var deleteResponse = proxy.OrdersDelete(orderId);

            if (deleteResponse.Errors != null && deleteResponse.Errors.Count > 0)
            {
                MessageBox.Show("Nem sikerült törölni a rendelést: " + string.Join(", ", deleteResponse.Errors));
                return;
            }

            originalOrder.Bvin = null; // nullázzuk, hogy új rendelésként jöjjön létre
            ApiResponse<Hotcakes.CommerceDTO.v1.Orders.OrderDTO> createResponse = proxy.OrdersCreate(originalOrder);

            if (createResponse.Errors.Any())
            {
                MessageBox.Show("Hiba a rendelés újralétrehozásakor: " + string.Join(", ", createResponse.Errors));
            }
            else
            {
                orderId = createResponse.Content.Bvin;
                MessageBox.Show("Rendelés sikeresen frissítve új létrehozással: " + createResponse.Content.Bvin);
            }

            LoadOrders(proxy);

        }

        private void LoadOrders(Api proxy)
        {
            var order = proxy.OrdersFind(orderId);

            items = order.Content.Items.Select(item => item.ProductName).ToList();
            quantity = order.Content.Items.Select(item => item.Quantity).ToList();
            if (items != null)
            {
                listBoxOrdered.DataSource = items;
            }
        }

        private void UpdateOrder_Remove()
        {
            Api proxy = ApiHivas();

            // 1. Rendelés lekérése
            var response = proxy.OrdersFind(orderId);
            var order = response.Content;

            // 2. Kiválasztott termék nevének lekérése a listboxból
            string selectedProductName = listBoxOrdered.Text?.ToString();
            if (string.IsNullOrEmpty(selectedProductName))
            {
                MessageBox.Show("Nincs kiválasztva termék a rendelésből.");
                return;
            }

            // 3. Termék megkeresése név alapján
            var allProducts = proxy.ProductsFindAll();
            var matchingProduct = allProducts.Content.FirstOrDefault(p => p.ProductName == selectedProductName);

            if (matchingProduct == null)
            {
                MessageBox.Show($"Nem található termék ilyen névvel: {selectedProductName}");
                return;
            }

            string productId = matchingProduct.Bvin;

            // 4. Meglévő tétel megkeresése és eltávolítása
            var existingItem = order.Items.FirstOrDefault(item => item.ProductId == productId);
            if (existingItem == null)
            {
                MessageBox.Show("A kiválasztott termék nem található ebben a rendelésben.");
                return;
            }

            order.Items.Remove(existingItem);

            // 5. Eredeti rendelés törlése
            var deleteResponse = proxy.OrdersDelete(order.Bvin);
            if (deleteResponse.Errors != null && deleteResponse.Errors.Count > 0)
            {
                MessageBox.Show("Hiba történt a rendelés törlése során: " + string.Join("\n", deleteResponse.Errors));
                return;
            }

            // 6. Bvin nullázása, hogy újként jöjjön létre
            order.Bvin = null;

            // 7. Új rendelés létrehozása módosított tételekkel
            var createResponse = proxy.OrdersCreate(order);
            if (createResponse.Errors != null && createResponse.Errors.Count > 0)
            {
                MessageBox.Show("Hiba a rendelés újralétrehozásakor: " + string.Join("\n", createResponse.Errors));
            }
            else
            {
                orderId = createResponse.Content.Bvin;
                MessageBox.Show("A termék sikeresen eltávolítva, a rendelés újra létrehozva!");
            }

            // 8. Rendelések újratöltése
            LoadOrders(proxy);
        }
    }
}
