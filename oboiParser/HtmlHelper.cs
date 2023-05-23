using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

namespace oboiParser
{
    public static class HtmlHelper
    {
        public static HtmlAgilityPack.HtmlDocument CreateDocument(this string HtmlPage)
        {
            HtmlAgilityPack.HtmlDocument document = new HtmlAgilityPack.HtmlDocument();

            document.LoadHtml(HtmlPage);
            return document;
        }
        public static  Product  ToProduct(this HtmlAgilityPack.HtmlDocument document) 
        { 
            var product = new Product();
            var root = document.DocumentNode;
            
            var pathNode = root.SelectNodes("//ol[@class='breadcrumbs']/li/a/span"); ;
            if(pathNode !=null)
            {
                product.Path = pathNode.Last().InnerText
                    .Replace("Главная", string.Empty)
                    .Replace("\t", string.Empty)
                    .Replace("\n", string.Empty)
                    .Replace("\r", string.Empty)
                    .Replace("▼", string.Empty)
                    .Trim();
            }
            
            //var imageNodes = root.SelectNodes("//div[class='product_image']/a"); //images_link
            var mainImageNode = root.SelectSingleNode("//img[@class='fn_img product_img']");
            if (mainImageNode !=null)
            {
                product.ImageUrls.Add(mainImageNode.Attributes["src"].Value);
            }
            var imageNodes = root.SelectNodes("//a[@class='images_link']");
            if(imageNodes !=null)
            {
                foreach (var imageNode in imageNodes)
                {
                    var imagHref = imageNode.Attributes["href"].Value;
                    product.ImageUrls.Add(imagHref.Replace("w.jp",".jp"));
                }
            }
            var headerNode = root.SelectSingleNode("//h1[@class='product_heading']");
            if(headerNode !=null)
            {
                product.Name = headerNode.InnerText.Trim();
            }
            var priceNode = root.SelectNodes("//div[@class='price  ']/span");
            if (priceNode != null)
            {
                product.Price = priceNode[0].InnerText.Trim();
                product.Сurrency = priceNode[1].InnerText.Trim();
            }
            else
            {
                var fn_priceNode = root.SelectSingleNode("//div[@class='fn_price']/span[@class='fn_old_price']");
                if(fn_priceNode != null)
                {
                    product.Price = fn_priceNode.InnerText.Trim();
                }
            }
            var tabtableNode = root.SelectSingleNode("//div[@class='tab_container comparison-mode']");
            if(tabtableNode != null) 
            {
                var tab1Node1 = tabtableNode.SelectSingleNode(".//div[@id='tab1']");
                if(tab1Node1 !=null)
                {
                    product.Appointment = tab1Node1.InnerHtml.Trim();
                }
                var tab1Node2 = tabtableNode.SelectSingleNode(".//div[@id='tab2']");
                if (tab1Node2 != null)
                {
                    product.Peculiarities = tab1Node2.InnerHtml.Trim();
                }
                
                var tab1Node3 = tabtableNode.SelectSingleNode(".//div[@id='tab3']");
                if (tab1Node3 != null)
                {
                    var tableNode3 = tab1Node3.SelectSingleNode(".//table");
                    if(tableNode3 !=null)
                    {
                        var borderAtt = tableNode3.Attributes["border"];
                        var classAtt = tableNode3.Attributes["class"];
                        
                        if  (classAtt != null && classAtt.Value == "tablepos")
                        {
                            var trNodes = tableNode3.SelectNodes(".//tbody/tr");
                            if (trNodes != null)
                            {
                                int double_count = 0;
                                string subCharacteristics = "";
                                for (int i = 1; i <= trNodes.Count - 1; i++)
                                {

                                    var tdNodes = trNodes[i].SelectNodes(".//td");

                                    if (tdNodes != null)
                                    {
                                        if (tdNodes.Count >= 2)
                                        {
                                            string key = (subCharacteristics + "." + tdNodes[0].InnerText.Trim()).TrimStart('.');
                                            string value = tdNodes[1].InnerText.Replace("\r\n"," ").Trim();
                                            if(product.Characteristics.ContainsKey(key))
                                            {
                                                key = key + $"_{double_count}";
                                                double_count++;
                                            }
                                            else
                                            {
                                                product.Characteristics.Add(key, value);
                                            }
                                           
                                        }
                                        else
                                        {
                                            var colspanAttr = tdNodes[0].Attributes["colspan"];
                                            if (colspanAttr == null)
                                                continue;
                                            if (colspanAttr.Value == "2")
                                            {
                                                subCharacteristics = tdNodes[0].InnerText.Trim(); ;


                                            }
                                            if (colspanAttr.Value == "3")
                                            {
                                                string key = subCharacteristics;
                                                string value = tdNodes[0].InnerText.Trim();
                                                product.Characteristics.Add(key, value);
                                            }


                                        }
                                    }
                                }
                               
                            }
                            else
                            {

                            }
                        }
                        else  if(borderAtt!=null && borderAtt.Value=="1")
                        {
                           product.IsTableCharacteristics = true;   
                        }
                        else if (classAtt != null && classAtt.Value == "brd")
                        {
                            product.IsTableCharacteristics = true;
                        }
                        else
                        {
                            product.IsTableCharacteristics = true;
                        }
                        if(product.IsTableCharacteristics)
                        {
                            product.OuterHtmlCharacteristics = tableNode3.OuterHtml;
                        }

                    }
                    else
                    {
                        
                    }


                }
                var tab1Node4 = tabtableNode.SelectSingleNode(".//div[@id='tab4']");
                if (tab1Node4 != null)
                {

                    product.ContentsOfDelivery = tab1Node4.InnerHtml.Trim();
                }
                var tab1Node5 = tabtableNode.SelectSingleNode(".//div[@id='tab5']");
                if (tab1Node5 != null)
                {

                    product.Documentation = tab1Node5.InnerHtml.Trim();
                }
            }

            

            return product;
        }
    }
}
