using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HamazonImportProduct
{
    public class Product
    {
        public int Mkt { get; set; }
        public string Barcode { get; set; }
        public string ProductName { get; set; }
        public string ProviderNewName { get; set; }
        public double WeightQuantity { get; set; }
        public bool Weighable { get; set; }
        public string Quantity { get; set; }
        public string PackageQuantity { get; set; }
        public int PackageQuantityInt { get; set; }
        public string UnitName { get; set; }
        public int UnitId { get; set; }
        public string BrandName { get; set; }
        public int BrandId { get; set; }
        public string ProviderName { get; set; }
        public int ProviderId { get; set; }
        public string GroupName { get; set; }
        public int GroupId { get; set; }
        public string Category1 { get; set; }
        public string Category2 { get; set; }
        public string Category3 { get; set; }

    }

    public class SimpleObject
    {
        public int Id { get; set; }
        public string Value { get; set; }
    }

    public class Category
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public Nullable<int> FatherId { get; set; }
        public Nullable<int> CategoryLevel { get; set; }
        public int Status { get; set; }
    }
}
