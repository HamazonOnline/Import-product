using HamazonImportProduct;
using Microsoft.VisualBasic.FileIO;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Xml.Linq;
using DAL;
using System.Text.RegularExpressions;
using System.Data.Entity.Core.Common.CommandTrees.ExpressionBuilder;

namespace HamazonImportProduct
{

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        List<Product> Products;
        List<Product> oldProducts;

        List<SimpleObject> brands;
        List<SimpleObject> providers;
        List<SimpleObject> unitMeasure;
        List<SimpleObject> families;
        List<Category> categories;
        public MainWindow()
        {
            brands = GetBrands();
            providers = GetProviders();
            unitMeasure = GetUnitMeasure();
            oldProducts = GetOldProducts();
            families = GetFamilies();
            categories = GetCategories();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // Set the license context here

            InitializeComponent();
            Products = ImportExcelFile();
            //foreach (var item in Products)
            //{
            //    if (item.Barcode.Length < 10)
            //    {
            //        string longBarcode = item.Barcode.PadLeft(10, '0'); //barcode length is 13 digits. always start with 729
            //        longBarcode = "729" + longBarcode;
            //        Console.WriteLine("barcode: " + item.Barcode + " long barcode: " + longBarcode);
            //        //find product in database per short/long barcode
            //        //if found - update parameters: status, 
            //        //if not found - add new produc
            //    }
            //}
        }

        private List<Category> GetCategories()
        {
            using(Entities context = new Entities())
            {
                List<Category> categories = context.p_category.Select(x => new Category
                {
                    Id = x.Id,
                    Name = x.Name,
                    FatherId = x.FatherId,
                    CategoryLevel = x.CategoryLevel,
                    Status = x.Active == true ? 1 : 0
                }).ToList();
                return categories;
            }
        }

        private List<SimpleObject> GetFamilies()
        {
            using(Entities context = new Entities())
            {
                List<SimpleObject> families = context.family.Select(x => new SimpleObject
                {
                    Id = x.Id,
                    Value = x.Name
                }).ToList();
                return families;
            }
        }

        private List<SimpleObject> GetUnitMeasure()
        {
            using (Entities context = new Entities())
            {
                List<SimpleObject> unitMeasure = context.p_unit_measure.Select(x => new SimpleObject
                {
                    Id = x.Id,
                    Value = x.Name
                }).ToList();
                return unitMeasure;
            }
        }

        public List<SimpleObject> GetBrands()
        {
            using(Entities context = new Entities())
            {
                List<SimpleObject> brands = context.p_brand.Select(x => new SimpleObject { 
                    Id = x.Id, 
                    Value = x.Name 
                }).ToList();
                return brands;
            }
        }

        public List<SimpleObject> GetProviders()
        {
            using(Entities context = new Entities())
            {
                List<SimpleObject> providers = context.p_provider.Select(x => new SimpleObject
                {
                    Id = x.Id, 
                    Value = x.Name 
                }).ToList();
                return providers;
            }
        }

        public List<Product> GetOldProducts()
        {
            using(Entities context = new Entities())
            {
                List<Product> oldProducts = context.product.Select(x => new Product
                {
                    Mkt = x.Mkt,
                    Barcode = x.Barcode,
                    ProductName = x.ProductName,
                    Quantity = x.WeightQuantity.ToString(),
                    UnitName = x.UnitMeasureId.ToString(),
                    PackageQuantity = x.PackageQuantity.ToString(),
                    BrandName = x.BrandId.ToString(),
                    ProviderName = x.ProviderId.ToString(),
                    GroupName = x.FamilyId.ToString(),

                    //Category2 = x.Category2,
                    //Category3 = x.Category3
                }).ToList();
                return oldProducts;
            }
        }

        public List<Product> ImportExcelFile()
        {
            string result="", brandResult = "", providerResult = "", unitResult = "", familyResult ="", categoryResult ="";
            List<Product> newProducts = new List<Product>();
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.DefaultExt = ".xlns";
            dlg.Filter = "Excel Files (*.xlsx)|*.xlsx";
            Nullable<bool> dlgResult = dlg.ShowDialog();
            if (dlgResult == true)
            {
                string filename = dlg.FileName;
                using (var package = new ExcelPackage(new FileInfo(filename)))
                {
                    foreach (var sheet in package.Workbook.Worksheets)//.Skip(2) skip the first 2 sheets
                    {
                        if (sheet.Dimension == null)
                        {
                            continue; // Skip empty sheets
                        }
                        // Get first row as string array
                        var headers = new string[sheet.Dimension.End.Column];
                        int emptyColumnIndex;
                        for (int i = 1; i <= sheet.Dimension.End.Column; i++)
                        {
                            headers[i - 1] = sheet.Cells[1, i].Text;
                        }
                        emptyColumnIndex = sheet.Dimension.End.Column;

                        //find column indexes
                        int barcodeIndex = Array.IndexOf(headers, "ברקוד");
                        int productNameIndex = Array.IndexOf(headers, "שם פריט");
                        int quantityIndex = Array.IndexOf(headers, "כמות");
                        int unitNameIndex = Array.IndexOf(headers, "יחידת מידה");
                        int packageQuantityIndex = Array.IndexOf(headers, "כמות במארז");
                        int brandNameIndex = Array.IndexOf(headers, "מותג");
                        int providerNameIndex = Array.IndexOf(headers, "ספק");
                        int groupNameIndex = Array.IndexOf(headers, "מקבץ");
                        int category1Index = Array.IndexOf(headers, "מחלקה");
                        int category2Index = Array.IndexOf(headers, "קבוצה");
                        int category3Index = Array.IndexOf(headers, "תת קבוצה");

                        if (barcodeIndex == -1)
                        {
                            continue;
                        }

                        for (int row = 2; row <= sheet.Dimension.Rows; row++)
                        {

                            result = ""; brandResult = ""; providerResult = ""; unitResult = ""; familyResult = "";
                            var p = new Product();

                            if (barcodeIndex == -1)
                            {
                                continue;
                            }
                            else
                            {
                                p.Barcode = sheet.Cells[row, barcodeIndex + 1].Text;
                                if (string.IsNullOrEmpty(p.Barcode))
                                {
                                    //write in excel "no barcode"
                                    //??? add with my mkt ???
                                    continue;
                                }
                            }
                            if (productNameIndex != -1) p.ProductName = sheet.Cells[row, productNameIndex + 1].Text;
                            if (quantityIndex != -1) p.Quantity = sheet.Cells[row, quantityIndex + 1].Text;
                            if (unitNameIndex != -1) p.UnitName = sheet.Cells[row, unitNameIndex + 1].Text;
                            if (packageQuantityIndex != -1) p.PackageQuantity = sheet.Cells[row, packageQuantityIndex + 1].Text;
                            if (brandNameIndex != -1) p.BrandName = sheet.Cells[row, brandNameIndex + 1].Text;
                            if (providerNameIndex != -1) p.ProviderName = sheet.Cells[row, providerNameIndex + 1].Text;
                            if (groupNameIndex != -1) p.GroupName = sheet.Cells[row, groupNameIndex + 1].Text;
                            if (category1Index != -1) p.Category1 = sheet.Cells[row, category1Index + 1].Text;
                            if (category2Index != -1) p.Category2 = sheet.Cells[row, category2Index + 1].Text;
                            if (category3Index != -1) p.Category3 = sheet.Cells[row, category3Index + 1].Text;

                            newProducts.Add(p);
                            string longBarcode = "";
                            if (p.Barcode.Length < 10)
                            {
                                longBarcode = p.Barcode.PadLeft(10, '0'); //barcode length is 13 digits. always start with 729
                                longBarcode = "729" + longBarcode;
                            }
                            else
                                longBarcode = p.Barcode;
                            string shortBarcode = p.Barcode;

                            List<Product> existProduct = oldProducts.Where(x => x.Barcode == shortBarcode || x.Barcode == longBarcode).ToList();

                            if (existProduct.Count > 0)
                            {
                                continue;
                            }


                            int id = GetProviderId(ref providerResult, p.ProviderName, providers);
                            if (id > 0) p.ProviderId = id;

                            id = GetBrandId(ref brandResult, p.BrandName, brands);
                            if (id > 0) p.BrandId = id;

                            id = GetFamilyId(ref familyResult, p.GroupName, families);
                            if (id > 0) p.GroupId = id;

                            if (p.Quantity == "שקיל")
                            {
                                p.Weighable = true;
                            }
                            else
                            {
                                p.Weighable = false;
                                bool b = double.TryParse(p.Quantity, out double qnt);
                                int i = GetObjectId(p.UnitName, unitMeasure);
                                if (b == true && i > 0)
                                {
                                    p.WeightQuantity = qnt;
                                    p.UnitId = i;
                                }
                                else
                                {
                                    unitResult = ($"{p.Quantity} {p.UnitName} שגיאה בפורמט כמות / ");
                                }
                            }
                               bool  bb = int.TryParse(p.PackageQuantity, out int pqnt);
                                if (bb == true)
                                {
                                    p.PackageQuantityInt = pqnt;
                                }


                                existProduct = oldProducts.Where(x => x.Barcode == shortBarcode || x.Barcode == longBarcode).ToList();

                                if (existProduct.Count == 0)
                                {
                                    p.Barcode = longBarcode;
                                    result = $"{p.ProductName} מוצר חדש";
                                    if (InsertProduct(p)  == null) 
                                        result += "/ נכשלה ההוספה";
                                }
                                else if (existProduct.Count == 1)
                                {
                                    result = "מוצר קיים";
                                    if (p.Barcode != existProduct[0].Barcode) result += longBarcode;
                                    if (!UpdateExistProductParameters(existProduct[0].Mkt, p)) result += "/ נכשל העדכון";
                                }
                                else if (existProduct.Count > 1)
                                {
                                    result = "ברקוד כפול";
                                }
                            
                            sheet.Cells[row, emptyColumnIndex + 1].Value = result;
                            sheet.Cells[row, emptyColumnIndex + 2].Value = providerResult;
                            sheet.Cells[row, emptyColumnIndex + 3].Value = brandResult;
                            sheet.Cells[row, emptyColumnIndex + 4].Value = unitResult;
                            sheet.Cells[row, emptyColumnIndex + 5].Value = familyResult;


                           //existProduct = oldProducts.Where(x => x.Barcode == shortBarcode || x.Barcode == longBarcode).ToList();
                            if(p.Mkt == null && p.Mkt == 0)
                            {
                                Console.WriteLine($"product {p.ProductName} already exist in database. barcode: {p.Barcode}");
                                continue;
                            }
                          

                            List<string> currentCategories= new List<string>();
                            if(string.IsNullOrEmpty(p.Category1))
                            {
                                categoryResult = "לא מופיעה מחלקה / ";
                                continue;
                            }
                            else
                            {
                                currentCategories.Add(p.Category1);
                                if (!string.IsNullOrEmpty(p.Category2))
                                {
                                    currentCategories.Add(p.Category2);
                                    if (!string.IsNullOrEmpty(p.Category3))
                                    {
                                        currentCategories.Add(p.Category3);
                                    }
                                }
                            }

                            int categoryIndex = 0;
                            int fatherId = 0;
                            int categoryId = 0;

                            categoryId = GetActiveLowCategory(currentCategories.Last());
                            if(categoryId == 0 )
                            {
                                foreach (var item in currentCategories)
                                {
                                    categoryId = GetOrInsertCategory(item, categoryIndex + 1, fatherId);
                                    if(categoryId == 0)
                                    {
                                        categoryResult = "שגיאה בהוספת קטגוריה / ";
                                        continue;
                                    }
                                    fatherId = categoryId;
                                    categoryIndex++;
                                }
                            }
                            // Connect product to category
                            if (categoryId > 0)
                            {
                                UpdateProductCategory(p.Mkt, categoryId);
                            }
                            continue;
                        }
                        Console.WriteLine("sheet: " + sheet.Name + "new product counter: " + newProducts.Count);
                        //save changes in excel
                        package.Save();
                        //xlApp.DisplayAlerts = false;
    //xlWorkBook.Save();

                        //xlWorkBook.Close();// Type.Missing, Type.Missing, Type.Missing);
                        //xlApp.Quit();
                    }           
                }
            }
            return newProducts;
        }

        private int GetActiveLowCategory(string categoryName)
        {
            using (Entities context = new Entities())
            {
                var existingCategory = context.p_category
                    .FirstOrDefault(c => c.Name == categoryName && c.Active == true);
                if(existingCategory != null)
                {
                    return existingCategory.Id;
                }
                else
                {
                    return 0;
                }
            }
        }

        private int GetOrInsertCategory(string categoryName, int categoryLevel, int fatherId)
        {
            using (Entities context = new Entities())
            {
                // Check if the category already exists
                var existingCategory = context.p_category
                    .FirstOrDefault(c => c.Name == categoryName);

                if (existingCategory != null)
                {
                    //active category
                    if(existingCategory.Active == false)
                    {
                        existingCategory.CategoryLevel = categoryLevel;
                        existingCategory.FatherId = fatherId;
                        existingCategory.Active = true;
                    context.SaveChanges();
                    }
                    
                    // Return the existing category's ID
                    return existingCategory.Id;
                }
                else
                {
                    // Create a new category
                    var newCategory = new p_category
                    {
                        Name = categoryName,
                        CategoryLevel = categoryLevel,
                        FatherId = fatherId,
                        Active = true // Assuming new categories are active by default
                    };

                    context.p_category.Add(newCategory);
                    context.SaveChanges();
                    categories.Add(new Category()
                    {
                        Id = newCategory.Id,
                        Name = newCategory.Name,
                        FatherId = newCategory.FatherId,
                        CategoryLevel = newCategory.CategoryLevel,
                        Status = newCategory.Active == true ? 1 : 0
                    });

                    // Return the new category's ID
                    return newCategory.Id;
                }
            }
        }

        private void ActiveCategoryStatus(int id)
        {
            using (Entities context = new Entities())
            {
                context.p_category.First(x => x.Id == id).Active = true;
                context.SaveChanges();
            }
        }

        private bool UpdateProductCategory(int mkt, int id)
        {
            using (Entities context = new Entities())
            {
                    context.product_category.Add(new product_category { ProductMkt = mkt, CategoryId = id });
                    try
                    {
                        context.SaveChanges();
                        return true;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);

                    }
                return false;
            }
        }

        private void updateSpecificParameterPerWrongResult()
        {
            //if (sheet.Cells[row, emptyColumnIndex + -5 + 4].Text.Contains("שגיאה בפורמט כמות"))
            //{
            //    bool b = double.TryParse(p.Quantity, out double qnt);
            //    int i = GetObjectId(p.UnitName, unitMeasure);
            //    if (b == true && i > 0)
            //    {
            //        p.WeightQuantity = qnt;
            //        p.UnitId = i;
            //    }
            //    else
            //    {
            //        sheet.Cells[row, emptyColumnIndex + -5 + 4].Value = ($"{p.Quantity} {p.UnitName} שגיאה בפורמט כמות / ");
            //        unitResult = ($"{p.Quantity} {p.UnitName} שגיאה בפורמט כמות / ");
            //    }
            //    List<Product> existProduct = oldProducts.Where(x => x.Barcode == shortBarcode || x.Barcode == longBarcode).ToList();
            //    foreach (var item in existProduct)
            //    {
            //        if (p.Barcode != item.Barcode) result += longBarcode;
            //        if (i != 0)
            //        {
            //            UpdateUnitMesure(item.Mkt, qnt, i);
            //            Console.WriteLine($"{row}-{emptyColumnIndex + 4 - 5}:   {p.Quantity} {p.UnitName} עודכנה כמות");
            //            sheet.Cells[row, emptyColumnIndex + 4 - 5].Value = "";
            //        }
            //        sheet.Cells[row, emptyColumnIndex + 4 - 5].Value = ($"{p.Quantity} {p.UnitName} שגיאה בפורמט כמות / ");
            //        continue;
            //    }
            //}
            //continue;
        }

        private void UpdateUnitMesure(int mkt, double qnt, int i)
        {
            using(Entities context= new Entities())
            {
                product p = context.product.First(x => x.Mkt == mkt);
                p.UnitMeasureId = i;
                p.WeightQuantity = qnt;
                //p.NewStatus = 1;
                context.SaveChanges();
            }
        }

        public int GetBrandId(ref string result ,string name, List<SimpleObject> brands)
        {
            if (name == "") return 0;
            int id = GetObjectId(name, brands);
            if(id ==0)
            {
                using(Entities context = new Entities())
                {
                    p_brand b = new p_brand();
                    b.Name = name;
                    try
                    {
                        context.p_brand.Add(b);
                        context.SaveChanges();
                        result = "הוספת מותג " + name;
                        id = b.Id;
                        brands.Add(new SimpleObject { Id = id, Value = name });
                    }
                    catch(Exception ex)
                    {
                        result = "שגיאה בהוספת מותג " + name;
                    }
                }
            }
            return id;
        }

        public int GetFamilyId(ref string result, string name, List<SimpleObject> families)
        {
            if (name == "") return 0;
            int id = GetObjectId(name, families);
            if (id == 0)
            {
                using (Entities context = new Entities())
                {
                    family f = new family();
                    f.Name = name;
                    try
                    {
                        context.family.Add(f);
                        context.SaveChanges();
                        result = "הוספת  מקבץ " + name;
                        id = f.Id;
                        families.Add(new SimpleObject { Id = id, Value = name });
                    }
                    catch (Exception ex)
                    {
                        result = "שגיאה בהוספת מקבץ " + name;
                    }
                }
            }
            return id;
        }

        public int GetProviderId(ref string result, string name, List<SimpleObject> providers)
        {
            if (name == "") return 0;
            int id = GetObjectId(name, providers);
            if (id == 0)
            {
                using (Entities context = new Entities())
                {
                    p_provider p = new p_provider();
                    p.Name = name;
                    try
                    {
                        context.p_provider.Add(p);
                        context.SaveChanges();
                        result = "הוספת ספק " + name;
                        id = p.Id;
                        providers.Add(new SimpleObject { Id = id, Value = name });
                    }
                    catch (Exception ex)
                    {
                        result = "שגיאה בהוספת ספק " + name;
                    }
                }
            }
            return id;
        }

        public int GetObjectId (string name, List<SimpleObject> list)
        {
            SimpleObject obj = list.FirstOrDefault(x => x.Value == name);
            if (obj != null)
            {
                return obj.Id;
            }
            else
            {
                return 0;
            }
        }

        private Product InsertProduct(Product p)
        {
            using (Entities context = new Entities())
            {
                product pro = new product();
                pro.Barcode = p.Barcode;
                pro.Weighable = p.Weighable;
                pro.ProductName = p.ProductName;
                pro.WeightQuantity = p.WeightQuantity;
                if(p.UnitId > 0) pro.UnitMeasureId = p.UnitId;
                if (p.PackageQuantityInt > 0) pro.PackageQuantity = p.PackageQuantityInt;
                if(p.BrandId > 0) pro.BrandId = p.BrandId;
                if(p.ProviderId > 0) pro.ProviderId = p.ProviderId;
                pro.NewStatus = 2;
                try
                {
                    context.product.Add(pro);
                    context.SaveChanges();
                    p.Mkt = pro.Mkt;
                    return p;
                }
                catch (Exception ex)
                {
                    Console.WriteLine("שגיאה בהוספת מוצר " + ex.Message);
                    return null;
                }
            }
        }

        public bool UpdateExistProductParameters(int mkt, Product p)
        {
            using(Entities context = new Entities())
            {
                product pro = context.product.First(x => x.Mkt == mkt);
                
                pro.Weighable = p.Weighable;
                pro.ProductName = p.ProductName;
                pro.WeightQuantity = p.WeightQuantity;
                if (p.UnitId > 0) pro.UnitMeasureId = p.UnitId;
                if (p.PackageQuantityInt > 0) pro.PackageQuantity = p.PackageQuantityInt;
                if (p.BrandId > 0) pro.BrandId = p.BrandId;
                if (p.ProviderId > 0) pro.ProviderId = p.ProviderId;
                pro.NewStatus = 1;
                try
                {
                    context.SaveChanges();
                    return true;
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    return false;
                }
            }
        }

        public List<Product> ImportCsvFile()
        {
            List<Product> newProducts = new List<Product>();
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            //dlg.DefaultExt = ".xlns";
            //dlg.Filter = "Excel Files (*.xlsx)|*.xlsx";
            dlg.DefaultExt = ".csv";
            dlg.Filter = "CSV Files (*.csv)|*.csv";
            Nullable<bool> result = dlg.ShowDialog();
            if (result == true)
            {
                // Open document
                string filename = dlg.FileName;
                //read csv file
                string[] lines = File.ReadAllLines(filename, System.Text.Encoding.GetEncoding("windows-1255"));
                string[] headers = lines[0].Split(',');
                
                lines = lines.Skip(1).ToArray(); //remove headers
                                                 //convert to list of newProductDTO

                foreach (var line in lines)
                {
                    TextFieldParser parser = new TextFieldParser(new StringReader(line));
                    parser.HasFieldsEnclosedInQuotes = true;
                    parser.SetDelimiters(",");
                    string[] data = null;
                    while (!parser.EndOfData)
                    {
                        data = parser.ReadFields();
                    }
                    parser.Close();

                    string[] values = line.Split(',');
                    Product p = new Product();

                    int barcodeIndex = Array.IndexOf(headers, "ברקוד");
                    int productNameIndex = Array.IndexOf(headers, "שם מוצר");
                    int quantityIndex = Array.IndexOf(headers, "כמות");
                    int unitNameIndex = Array.IndexOf(headers, "יחידת מידה");
                    int packageQuantityIndex = Array.IndexOf(headers, "כמות במארז");
                    int brandNameIndex = Array.IndexOf(headers, "מותג");
                    int providerNameIndex = Array.IndexOf(headers, "ספק");
                    int groupNameIndex = Array.IndexOf(headers, "מקבץ");
                    int category1Index = Array.IndexOf(headers, "מחלקה");
                    int category2Index = Array.IndexOf(headers, "קבוצה");
                    int category3Index = Array.IndexOf(headers, "תת קבוצה");

                    if (barcodeIndex != -1) p.Barcode = data[barcodeIndex];
                    p.ProductName = data[productNameIndex];
                    p.Quantity = data[quantityIndex];
                    p.UnitName = data[unitNameIndex];
                    p.PackageQuantity = data[packageQuantityIndex];
                    p.BrandName = data[brandNameIndex];
                    p.ProviderName = data[providerNameIndex];
                    p.GroupName = data[groupNameIndex];
                    p.Category1 = data[category1Index];
                    p.Category2 = data[category2Index];
                    p.Category3 = data[category3Index];

                    newProducts.Add(p);

                    //    newProducts.Add(n);

                    //}
                    //if (newProducts.Count == 0)
                    //{
                    //    MessageBox.Show("אין מוצרים חדשים לייבא");
                    //    return;
                    //}
                }
            }
            return newProducts;
        }
    }
}