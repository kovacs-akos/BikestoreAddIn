using Bikestore.BikestoreModels;
using Microsoft.EntityFrameworkCore;

namespace Bikestore
{
    public partial class Form1 : Form
    {
        private readonly SeBikestoreContext _context = new();

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dataGridView1.DataSource = _context.Orders
            .Include(o => o.CustomerFkNavigation)
            .Include(o => o.StoreFkNavigation)
            .OrderBy(o => o.OrderSk)
            .Select(o => new
            {
                o.OrderSk,
                Customer = o.CustomerFkNavigation.FirstName + " " + o.CustomerFkNavigation.LastName,
                o.CustomerFkNavigation.City,
                o.OrderDate,
                o.OrderStatus,
                Store = o.StoreFkNavigation.StoreName
            })
            .ToList();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            new ExcelExport().CreateExcel();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string path = new ClosedXmlExport().Export();
            MessageBox.Show($"Mentve: {path}");
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(path) { UseShellExecute = true });
        }
    }
}
