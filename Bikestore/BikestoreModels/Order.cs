using System;
using System.Collections.Generic;

namespace Bikestore.BikestoreModels;

public partial class Order
{
    public int OrderSk { get; set; }

    public int? CustomerFk { get; set; }

    public byte OrderStatus { get; set; }

    public DateOnly OrderDate { get; set; }

    public DateOnly RequiredDate { get; set; }

    public DateOnly? ShippedDate { get; set; }

    public int StoreFk { get; set; }

    public int StaffFk { get; set; }

    public virtual Customer? CustomerFkNavigation { get; set; }

    public virtual ICollection<OrderItem> OrderItems { get; set; } = new List<OrderItem>();

    public virtual Staff StaffFkNavigation { get; set; } = null!;

    public virtual Store StoreFkNavigation { get; set; } = null!;
}
