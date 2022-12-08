using System.Collections.Generic;

namespace GoogleSheetAPI
{
    public static class ProductsMapper
    {
        public static List<Product> MapFromRangeData(IList<IList<object>> values)
        {
            var products = new List<Product>();

            foreach (var value in values)
            {
                Product product = new()
                {
                    Id = value[0].ToString(),
                    Name = value[1].ToString(),
                    Category = value[2].ToString(),
                    Price = value[3].ToString()
                };

                products.Add(product);
            }

            return products;
        }

        public static IList<IList<object>> MapToRangeData(Product product)
        {
            var objectList = new List<object>() 
                                    { 
                                        product.Id, 
                                        product.Name, 
                                        product.Category, 
                                        product.Price 
                                    };
            var rangeData = new List<IList<object>> { objectList };
            return rangeData;
        }
    }
}
