using System;
using System.Collections.Generic;

namespace Bikestore.BikestoreModels;

public partial class Stock
{
    public int StoreSk { get; set; }

    public int ProductFk { get; set; }

    public int? Quantity { get; set; }

    public virtual Product ProductFkNavigation { get; set; } = null!;

    public virtual Store StoreSkNavigation { get; set; } = null!;
}
