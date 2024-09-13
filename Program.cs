using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;
using Newtonsoft.Json;
using System.Linq;
using System.Net;
using System.Net.Sockets;

class Program
{
    static async Task Main()
    {
        // Define the path to the Excel file
        string filePath = @"..\..\..\stock.xlsx";

        int completedProduct = 0;
        // Initialise a list to hold the data
        List<StockItem> stockItems = new List<StockItem>();
        // Configure EPPlus to read the Excel file
        FileInfo fileInfo = new FileInfo(filePath);

        string currentBrand = null;

        Dictionary<string, List<StockItem>> brandGroupedStockItems = new Dictionary<string, List<StockItem>>();

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // Add for EPPlus 5 or later

        using (ExcelPackage package = new ExcelPackage(fileInfo))
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Get the first worksheet

            // Iterate through the rows, starting from the second row if there's a header
            for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
            {
                string column1Value = worksheet.Cells[row, 1].Text; // First column
                string column2Value = worksheet.Cells[row, 2].Text; // Second column
                string hyperlink = worksheet.Cells[row, 1].Hyperlink?.AbsoluteUri ?? string.Empty; // Extract hyperlink if available
                
        

                if (!string.IsNullOrWhiteSpace(column1Value) && string.IsNullOrWhiteSpace(column2Value))
                {
                    // This is a brand row
                    currentBrand = column1Value;
                    // Initialize the list for new brand if it does not exist
                    if (!brandGroupedStockItems.ContainsKey(currentBrand))
                    {
                        brandGroupedStockItems[currentBrand] = new List<StockItem> ();
                    }
                    continue;
                }


                // Only continue if at least one of the first two columns is not empty
                if (string.IsNullOrWhiteSpace(column1Value) && string.IsNullOrWhiteSpace(column2Value))
                {
                    continue;
                }

                // https://images.stockx.com/360/adidas-Yeezy-450-Dark-Slate/Images/adidas-Yeezy-450-Dark-Slate/Lv2/img01.jpg?auto=format%2Ccompress&w=576&dpr=1&updated_at=1640193016&h=384&q=57
                // https://images.stockx.com/360/adidas-yeezy-450-dark-slate/Images/adidas-yeezy-450-dark-slate/Lv2/img01.jpg?auto=format%2Ccompress&w=480&dpr=1&updated_at=

                // Create a StockItem object and add it to the list
                if (currentBrand != null && (!string.IsNullOrWhiteSpace(column1Value) || !string.IsNullOrWhiteSpace(column2Value)))
                {
                    string shoeLink = ExtractFormattedProductName(hyperlink);
                    string shoeImageLink = $"https://images.stockx.com/360/{shoeLink}/Images/{shoeLink}/Lv2/img01.jpg?auto=format%2Ccompress&w=480&dpr=1&updated_at=";
                    string essentialsFOGLink = $"https://images.stockx.com/images/vertical/Fear-Of-God-Essentials-Hoodie-Black_1.jpg?fit=fill&bg=FFFFFF&w=396&h=504&auto=format%2Ccompress&dpr=1&q=57";
                    // https://images.stockx.com/images/Fear-of-God-Essentials-1977-Crewneck-Iron.jpg?fit=fill&bg=FFFFFF&w=576&h=384&auto=format%2Ccompress&dpr=1&trim=color&updated_at=1646952565&q=57

                    bool success = await ValidateUrlWithHttpClient(shoeImageLink);

                   // Console.WriteLine($"{success}");

                    // Create a StockItem object and add it to the list for the current brand
                    StockItem stockItem = new StockItem
                    {
                        name = column1Value,
                        ProductSize = column2Value,
                        productUrl = hyperlink,
                        imageUrl = shoeImageLink,
                        ImageLinkWorks = success
                    };

                    completedProduct++;
                    DisplayProgressBar(completedProduct, worksheet.Cells.Count());
                    

                    brandGroupedStockItems[currentBrand].Add(stockItem);
                }

               //Console.WriteLine($"Row {row}: Column1 = {column1Value}, Column2 = {column2Value}, Hyperlink = {hyperlink}");
            }
        }

        // Convert the list of stock items to JSON
        string jsonResult = JsonConvert.SerializeObject(brandGroupedStockItems, Formatting.Indented);

        string outputPath = @"..\..\..\stock_items.txt"; // Specify the path to your desired output file
        Console.WriteLine(jsonResult);

        File.WriteAllText(outputPath, jsonResult);
    }



      public static void DisplayProgressBar(int current, int total)
        {
            Console.CursorLeft = 0;
            Console.Write("[");
            int progressWidth = Console.WindowWidth - 30; // Adjust width if needed
            int position = (int)((double)current / total * progressWidth);

            Console.Write(new string('=', position));
            Console.Write(new string(' ', progressWidth - position));
            Console.Write($"] {current * 100 / total}% ({current}/{total})");
        }


    public static async Task<bool> ValidateUrlWithHttpClient(string url)
    {
        using var client = new HttpClient();
        try
        {
            var response = await client.SendAsync(new HttpRequestMessage(HttpMethod.Head, url));

            return response.IsSuccessStatusCode;
        }
        catch (HttpRequestException e)
            when (e.InnerException is SocketException
            { SocketErrorCode: SocketError.HostNotFound })
        {
            return false;
        }
        catch (HttpRequestException e)
            when (e.StatusCode.HasValue && (int)e.StatusCode.Value > 500)
        {
            return true;
        }
    }




    static string ExtractFormattedProductName(string url)
    {
        // Only extract the path segment of the URL
        Uri uri;
        if (Uri.TryCreate(url, UriKind.Absolute, out uri))
        {
            string[] segments = uri.AbsolutePath.Trim('/').Split('/');
            if (segments.Length > 0)
            {
                string productName = segments[0]; // Get the product name segment
                if (productName.StartsWith("fear-of-god") || productName.StartsWith("nike") )
                {
                    return CapitalizeEachWord(productName);
                }
                else if (productName.StartsWith("adidas-yeezy"))
                {
                    return CapitalizeAfterFirstWord(productName);
                }


                // Add additional rules if needed
                return productName;
            }
        }
        return string.Empty;
    }

    static string CapitalizeEachWord(string input)
    {
        var words = input.Split('-');
        for (int i = 0; i < words.Length; i++)
        {
            words[i] = char.ToUpper(words[i][0]) + words[i].Substring(1);
        }
        return string.Join('-', words);
    }

    static string CapitalizeAfterFirstWord(string input)
    {
        var words = input.Split('-');
        for (int i = 1; i < words.Length; i++)
        {
            if (char.IsLower(words[i][0]))
            {
                words[i] = char.ToUpper(words[i][0]) + words[i].Substring(1);
            }
        }
        return string.Join('-', words);
    }
}
    //static string ExtractProductNameFromUrl(string url)
    //{
    //    // Example URL format: https://stockx.com/adidas-yeezy-450-stone-flax?size=10
    //    Uri uri;
    //    if (Uri.TryCreate(url, UriKind.Absolute, out uri))
    //    {
    //        // PathAndQuery gives the path part without the domain
    //        var pathSegments = uri.PathAndQuery.Split('/');
    //        if (pathSegments.Length > 1)
    //        {
    //            string productNameWithQuery = pathSegments[1];
    //            int queryIndex = productNameWithQuery.IndexOf('?');
    //            if (queryIndex > 0)
    //            {
    //                return productNameWithQuery[..queryIndex]; // Extract up to but not including '?'
    //            }
    //            return productNameWithQuery;
    //        }
    //    }
    //    return string.Empty;
    // }
    //}
// https://images.stockx.com/360/SHOENAME/Images/SHOENAME/Lv2/img01.jpg?auto=format%2Ccompress&w=480&dpr=1&updated_at=

public class StockItem
{
    public string? name { get; set; }
    public string? ProductSize { get; set; }
    public string? productUrl { get; set; }
    public string? imageUrl { get; set; }
    public bool? ImageLinkWorks { get; set; }
}

public class StockByBrand
{
    
}
