using Bikestore.BikestoreModels;
using ClosedXML.Excel;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Text;

namespace Bikestore
{
    public class ClosedXmlExport
    {
        private readonly SeBikestoreContext _context = new();

        public string Export()
        {
            var orders = _context.Orders
                .Include(o => o.CustomerFkNavigation)
                .Include(o => o.StoreFkNavigation)
                .Include(o => o.StaffFkNavigation)
                .Include(o => o.OrderItems)
                .OrderBy(o => o.OrderSk)
                .ToList();

            using var workbook = new XLWorkbook();
            var ws = workbook.Worksheets.Add("Orders");

            string[] fejlecek =
            {
            "ID", "Ügyfél", "Város", "Állam", "Rendelés dátuma",
            "Szállítás dátuma", "Státusz", "Bolt", "Eladó", "Végösszeg"
        };
            for (int i = 0; i < fejlecek.Length; i++)
                ws.Cell(1, i + 1).Value = fejlecek[i];

            for (int i = 0; i < orders.Count; i++)
            {
                var o = orders[i];
                int sor = i + 2;
                ws.Cell(sor, 1).Value = o.OrderSk;
                ws.Cell(sor, 2).Value = $"{o.CustomerFkNavigation.FirstName} {o.CustomerFkNavigation.LastName}";
                ws.Cell(sor, 3).Value = o.CustomerFkNavigation.City;
                ws.Cell(sor, 4).Value = o.CustomerFkNavigation.State;
                ws.Cell(sor, 5).Value = o.OrderDate.ToDateTime(TimeOnly.MinValue);   // DateOnly → DateTime, marad dátum!
                if (o.ShippedDate is { } szallitva)
                    ws.Cell(sor, 6).Value = szallitva.ToDateTime(TimeOnly.MinValue);
                ws.Cell(sor, 7).Value = o.OrderStatus;                                // marad szám
                ws.Cell(sor, 8).Value = o.StoreFkNavigation.StoreName;
                ws.Cell(sor, 9).Value = $"{o.StaffFkNavigation.FirstName} {o.StaffFkNavigation.LastName}";
                ws.Cell(sor, 10).Value = o.OrderItems.Sum(t => t.Quantity * (t.ListPrice ?? 0) * (1 - (t.Discount ?? 0))); // decimal, marad szám
            }

            // Formázás – itt Excel számformátumot adunk, nem ToString-et
            var fejlecRange = ws.Range(1, 1, 1, fejlecek.Length);
            fejlecRange.Style.Font.Bold = true;
            fejlecRange.Style.Fill.BackgroundColor = XLColor.Fuchsia;
            fejlecRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            fejlecRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thick;

            ws.Column(5).Style.DateFormat.Format = "yyyy-MM-dd";
            ws.Column(6).Style.DateFormat.Format = "yyyy-MM-dd";
            ws.Column(10).Style.NumberFormat.Format = "#,##0.00 \"$\"";
            ws.Columns().AdjustToContents();

            // Egy összegsor – a képletet az Excel számolja
            int utolso = orders.Count + 1;
            ws.Cell(utolso + 1, 9).Value = "Összesen:";
            ws.Cell(utolso + 1, 10).FormulaA1 = $"SUM(J2:J{utolso})";

            string path = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                $"Orders_{DateTime.Now:yyyyMMdd_HHmm}.xlsx");
            workbook.SaveAs(path);
            return path;
        }
    }
}
