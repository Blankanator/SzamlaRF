using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Rendeleskezelo;
using System.IO;
using System.Windows.Forms;
// Removed: using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace UnitTestExample.Test
{
    internal class teszt
    {
        [Test]
        public void Test_GenerateXmlFile_CreatesFileWithContent()
        {
            // Arrange
            var form = new MainForm();

            // DataGridView el�r�se publikus property-n kereszt�l
            var grid = form.OrdersGrid;

            // Oszlopok hozz�ad�sa, ha nincsenek (teszt k�zben nem automatikus)
            if (grid.Columns.Count == 0)
            {
                grid.Columns.Add(new DataGridViewCheckBoxColumn() { Name = "checkBox" });
                grid.Columns.Add("orderId", "OrderId");
                grid.Columns.Add("customerName", "Customer Name");
                grid.Columns.Add("email", "Email");
                grid.Columns.Add("orderDate", "Order Date");
                grid.Columns.Add("shippingAddress", "Shipping Address");
                grid.Columns.Add("orderTotal", "Order Total");
            }

            // Sor hozz�ad�sa, bejel�lve (checkbox = true)
            grid.Rows.Add(true, "123", "Teszt Elek", "teszt@valami.hu", "2025-01-01", "Teszt utca 1", "1500");

            // Teszt f�jln�v
            string filePath = "teszt.xml";
            if (File.Exists(filePath))
                File.Delete(filePath);

            // Act
            form.GenerateXMLToFile(filePath);

            // Assert
            Assert.IsTrue(File.Exists(filePath), "Nem j�tt l�tre az XML f�jl.");
            string content = File.ReadAllText(filePath);
            Assert.IsTrue(content.Contains("Teszt Elek"), "A f�jl nem tartalmazza a v�rt adatokat.");
        }
    }
}
