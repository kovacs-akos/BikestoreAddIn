using Bikestore.BikestoreModels;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;


namespace Bikestore
{
    internal class ExcelExport
    {
        private readonly SeBikestoreContext _context = new();
        private Excel.Application xlApp = null!;
        private Excel.Workbook xlWB = null!;
        private Excel.Worksheet xlOrdersSheet = null!;

        private readonly string[] fejlecek =
    {
        "ID", "Ügyfél", "Város", "Állam", "Rendelés dátuma",
        "Szállítás dátuma", "Státusz", "Bolt", "Eladó", "Végösszeg"
    };

        // order_status: 1 = Pending, 2 = Processing, 3 = Rejected, 4 = Completed
        private static string StatusText(byte status) => status switch
        {
            1 => "Függőben",
            2 => "Feldolgozás alatt",
            3 => "Elutasítva",
            4 => "Teljesítve",
            _ => status.ToString()
        };

        public void CreateExcel()
        {
            try
            {
                xlApp = new Excel.Application();
                xlWB = xlApp.Workbooks.Add(Missing.Value);
                xlOrdersSheet = (Excel.Worksheet)xlWB.ActiveSheet;
                xlOrdersSheet.Name = "Orders";

                CreateTable();
                FormatTable();

                xlApp.Visible = true;
                xlApp.UserControl = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Hiba: {ex.Message}", "Excel export");
                xlWB?.Close(false, Type.Missing, Type.Missing);
                xlApp?.Quit();
                xlWB = null!;
                xlApp = null!;
            }
        }
        
        
        private void CreateTable()
        {
            // Fejléc: tömb 0-tól, Excel cella 1-től indexel!
            for (int i = 0; i < fejlecek.Length; i++)
                xlOrdersSheet.Cells[1, i + 1] = fejlecek[i];

            var orders = _context.Orders
                .Include(o => o.CustomerFkNavigation)
                .Include(o => o.StoreFkNavigation)
                .Include(o => o.StaffFkNavigation)
                .Include(o => o.OrderItems)
                .OrderBy(o => o.OrderSk)
                .ToList();

            object[,] adatTomb = new object[orders.Count, fejlecek.Length];

            for (int i = 0; i < orders.Count; i++)
            {
                var o = orders[i];
                decimal vegosszeg = o.OrderItems.Sum(t => t.Quantity * (t.ListPrice ?? 0) * (1 - (t.Discount ?? 0)));

                adatTomb[i, 0] = o.OrderSk;
                adatTomb[i, 1] = $"{o.CustomerFkNavigation.FirstName} {o.CustomerFkNavigation.LastName}";
                adatTomb[i, 2] = o.CustomerFkNavigation.City;
                adatTomb[i, 3] = o.CustomerFkNavigation.State;
                adatTomb[i, 4] = o.OrderDate.ToString("yyyy-MM-dd");
                adatTomb[i, 5] = o.ShippedDate?.ToString("yyyy-MM-dd") ?? "";
                adatTomb[i, 6] = StatusText(o.OrderStatus);
                adatTomb[i, 7] = o.StoreFkNavigation.StoreName;
                adatTomb[i, 8] = $"{o.StaffFkNavigation.FirstName} {o.StaffFkNavigation.LastName}";
                adatTomb[i, 9] = vegosszeg.ToString("C2");
            }

            int sorokSzama = adatTomb.GetLength(0);
            int oszlopokSzama = adatTomb.GetLength(1);

            Excel.Range adatRange = xlOrdersSheet
                .get_Range("A2", Type.Missing)
                .get_Resize(sorokSzama, oszlopokSzama);
            adatRange.Value2 = adatTomb;
            adatRange.Columns.AutoFit();
        }


        private void FormatTable()
        {
            int utolsoSor = xlOrdersSheet.UsedRange.Rows.Count;
            int utolsoOszlop = fejlecek.Length;

            Excel.Range fejlecRange = xlOrdersSheet
                .get_Range("A1", Type.Missing)
                .get_Resize(1, utolsoOszlop);
            fejlecRange.Font.Bold = true;
            fejlecRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            fejlecRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            fejlecRange.RowHeight = 40;
            fejlecRange.Interior.Color = Color.Fuchsia;
            fejlecRange.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick);

            // Gyakorló feladatok:
            // 1. az egész tábla körül vastag szegély (A1 → utolsoSor x utolsoOszlop)
            // 2. első oszlop félkövér + Color.LightYellow háttér (A2-től)
            // 3. utolsó oszlop Color.LightGreen háttér
            // 4. utolsó oszlop NumberFormat = "#,##0.00" (ehhez a vegosszeg-et számként kell átadni: (double)vegosszeg, ne ToString-gel!)

            fejlecRange.EntireColumn.AutoFit();
        }
    }
}

