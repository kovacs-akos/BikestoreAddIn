using System;
using System.Collections.Generic;

namespace Bikestore.BikestoreModels;

public partial class Product
{
    public int ProductSk { get; set; }

    public string ProductName { get; set; } = null!;

    public int BrandId { get; set; }

    public int CategoryFk { get; set; }

    public short ModelYear { get; set; }

    public decimal ListPrice { get; set; }

    public virtual Brand Brand { get; set; } = null!;

    public virtual Category CategoryFkNavigation { get; set; } = null!;

    public virtual ICollection<OrderItem> OrderItems { get; set; } = new List<OrderItem>();

    public virtual ICollection<Stock> Stocks { get; set; } = new List<Stock>();
}
