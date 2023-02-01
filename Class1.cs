using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Threading;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using OpenQA.Selenium.Support.UI;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ShoppingCart_Excel
{
    public class Class1 : Input_Class
    {
        IWebDriver driver;

        public void startBrowser()
        {
            driver = new ChromeDriver();
            driver.Manage().Window.Maximize();
            driver.Url = "http://tutorialsninja.com/demo/index.php?route=product/category&path=20";
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
        }

        //to gather all the product names present on the site and write it into excel
        public void getProductName()
        {
            xlapp.Visible = true;
            try
            {
                xlRange.Cells[1, 1] = "Product Name";
                IList<IWebElement> prd_name = driver.FindElements(By.XPath("//div[@class='caption']/h4"));
                Thread.Sleep(2000);
                for (int i = 0; i < prd_name.Count; i++)
                {
                    xlRange.Cells[i + 2, 1] = prd_name[i].Text;
                }
                xlwb.SaveAs(filepath);
            }
            catch (Exception exHandle)
            {
                Console.WriteLine("Exception: " + exHandle.Message);
                Console.ReadLine();
            }
        }

        //Gathering data like Price, Discounted price and product code and write it into excel
        public void GetBasicDetails()
        {
            xlwb = xlapp.Workbooks.Open(filepath);
            xlws = xlwb.Worksheets[1];
            xlRange = xlws.UsedRange;
            rowCount = xlRange.Rows.Count;
            colCount = xlRange.Columns.Count;
            xlRange.Cells[1, 2] = "Product Price";
            xlRange.Cells[1, 3] = "Product Discounted Price";
            xlRange.Cells[1, 5] = "Product Code";
            for (int i = 0; i <= rowCount - 2; i++)
            {
                driver.FindElement(By.XPath("//li/a[text()='Desktops']")).Click();
                driver.FindElement(By.XPath("//*[text()='Show All Desktops']")).Click();
                string p = xlRange.Cells[i + 2, 1].Value2.ToString();
                driver.FindElement(By.XPath("//h4/a[text()='" + p + "']")).Click();
                IWebElement price_pull = driver.FindElement(By.XPath("//div[@class='col-sm-4']/ul[2]/li[1]"));
                price.Add(price_pull.Text.ToString());
                xlRange.Cells[i + 2, 2] = price[i];
                if (p == "Apple Cinema 30\"" || p == "Canon EOS 5D")
                {
                    IWebElement price_pull2 = driver.FindElement(By.XPath("//div[@class='col-sm-4']/ul[2]/li[2]"));
                    discounted_price.Add(price_pull2.Text.ToString());
                    xlRange.Cells[i + 2, 3] = discounted_price[i];
                }
                else
                {
                    xlRange.Cells[i + 2, 3] = "NA";
                }
                IWebElement prd_code = driver.FindElement(By.XPath("//*[contains(text(), 'Product Code')]"));
                product_code.Add(prd_code.Text.ToString());
                xlRange.Cells[i + 2, 5] = product_code[i];
            }
            final_price();
            final_productcode();
            quantity_prd();
        }

        public void final_price()
        {
            xlRange.Cells[1, 4] = "Final Price";
            for (int i=0; i<=rowCount-2; i++)
            {
                string f = xlRange.Cells[i + 2, 3].Value2.ToString();
                if (f=="NA")
                {
                    xlRange.Cells[i + 2, 4] = xlRange.Cells[i + 2, 2];
                }
                else
                {
                    xlRange.Cells[i + 2, 4] = xlRange.Cells[i + 2, 3];
                }
            }
        }

        public void final_productcode()
        {
            xlRange.Cells[1, 6] = "Product Code";
            for(int i =0; i<=product_code.Count-1; i++)
            {
                string x = product_code[i].ToString();
                y = x.Split(':');
                xlRange.Cells[i+2, 6] = y[1].TrimStart();
            }
            xlws.Columns["E"].Delete();
        }

        public void quantity_prd()
        {
            xlRange.Cells[1, 6] = "Quantity";
            for(int i=0; i<quant.Count; i++)
            {
                xlRange.Cells[i + 2, 6] = quant[i];
            }
        }

        //Getting data for test method one and using it to add products to cart
        public void AddProduct()
        {
            xlwb = xlapp.Workbooks.Open(filepath);
            xlws = xlwb.Worksheets[1];
            xlRange = xlws.UsedRange;
            rowCount = xlRange.Rows.Count;
            colCount = xlRange.Columns.Count;
            try
            {
                for (int i = 2; i <= rowCount; i++)
                {
                    for (int j = 1; j <= 1; j++)
                    {
                        driver.FindElement(By.XPath("//li/a[text()='Desktops']")).Click();
                        driver.FindElement(By.XPath("//*[text()='Show All Desktops']")).Click();
                        string excel_prdname = xlRange.Cells[i, j].Value2.ToString();
                        string excel_quant = xlRange.Cells[i, j + 5].Value2.ToString();

                        //if loop for Apple Cinema as it has some radio buttons and checkboxes which needs to be checked, right now its hardcoded
                        if (excel_prdname == "Apple Cinema 30\"")
                        {
                            driver.FindElement(By.XPath("//h4/a[text()='" + excel_prdname + "']")).Click();
                            driver.FindElement(By.XPath("//input[@value='7']")).Click();
                            driver.FindElement(By.XPath("//input[@value='9']")).Click();
                            SelectElement sel = new SelectElement(driver.FindElement(By.Id("input-option217")));
                            sel.SelectByText("Blue (+$3.60)");
                            driver.FindElement(By.Id("input-option209")).SendKeys("Testing");
                            IWebElement test = driver.FindElement(By.Id("button-upload222"));
                            test.Click();
                            Thread.Sleep(2000);
                            SendKeys.SendWait(@"C:\Users\v-ggaikwad\Documents\Tools and something\OIP.jpg");
                            SendKeys.SendWait("{Enter}");
                            Thread.Sleep(2000);
                            driver.SwitchTo().Alert().Accept();
                            driver.FindElement(By.Id("input-quantity")).Clear();
                            driver.FindElement(By.Id("input-quantity")).SendKeys(excel_quant);
                            driver.FindElement(By.XPath("//*[@id='tab-description']/ul[3]/li[4]")).Click();
                            driver.FindElement(By.XPath("//*[text()='Add to Cart']")).Click();
                        }
                        if (excel_prdname != "Apple Cinema 30\"" && excel_prdname!= "Canon EOS 5D")
                        {
                            driver.FindElement(By.XPath("//h4/a[text()='" + excel_prdname + "']")).Click();
                            try
                            {
                                IWebElement dropdw = driver.FindElement(By.XPath("//select[@class='form-control']"));
                                SelectElement s = new SelectElement(dropdw);
                                s.SelectByValue("13");
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.Message);
                            }
                            driver.FindElement(By.Id("input-quantity")).Clear();
                            driver.FindElement(By.Id("input-quantity")).SendKeys(excel_quant);
                            driver.FindElement(By.XPath("//button[text()='Add to Cart']")).Click();

                        }
                        if(excel_prdname== "Canon EOS 5D")
                        {
                            Console.WriteLine("Do Nothing");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                driver.Quit();
                xlwb.Close();
                Console.WriteLine(ex.Message);
            }
        }


        //Verify quantity present on cart with the input quantity
        public void verifyqauntity()
        {
            for(int i=1; i<rowCount-1; i++)
            {
                IWebElement Quantity = driver.FindElement(By.XPath("//tr["+i+"]/td[4]/div/input[@class='form-control']"));
                string q = Quantity.GetAttribute("value");
                site_quant.Add(q);
            }

            for(int k = 1; k <= site_quant.Count; k++) 
            {
                try
                {
                    IWebElement site_prdname = driver.FindElement(By.XPath("//div/table/tbody/tr["+k+"]/td[@class='text-left']/a"));
                    string prdname = site_prdname.Text;
                    for (int j = 2; j <= rowCount; j++)
                    {
                        string excel_prdname = xlRange.Cells[j, 1].Value2.ToString();
                        try
                        {
                            if (prdname == excel_prdname)
                            {
                                Assert.AreEqual(site_quant[k-1], xlRange.Cells[j, 6].Value2.ToString(), "Result Not Found");
                                break;
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
                
            }
        }

        //Verify unit price present on cart with excel price
        public void verifyUnitPrice()
        {
            for(int i=1; i< rowCount-1; i++)
            {
                IWebElement site_prdname = driver.FindElement(By.XPath("//div/table/tbody/tr[" + i + "]/td[@class='text-left']/a"));
                string prdname = site_prdname.Text;
                IWebElement site_uprice = driver.FindElement(By.XPath("//div/table/tbody/tr["+i+"]/td[5][@class='text-right']"));
                string expected_up = site_uprice.Text;
                for (int j=2; j<=rowCount; j++)
                {
                    string excel_prdname = xlRange.Cells[j, 1].Value2.ToString();
                    double excel_unitprice = xlRange.Cells[j, 4].Value2;
                    try
                    {
                        if(prdname == excel_prdname)
                        {
                            //if loop here since Apple cinema and product 8 containg dropdowns and radio buttons which adds up the price
                            if (prdname == "Apple Cinema 30\"")
                            {
                                double new_exup = excel_unitprice + 63.6;
                                string final_exup = new_exup.ToString("N2");
                                Assert.AreEqual("$" + final_exup, expected_up, "Result Not Found");
                                break;
                            }
                            if (prdname == "Product 8")
                            {
                                double new_exup = excel_unitprice + 12;
                                string final_exup = new_exup.ToString("N2");
                                Assert.AreEqual("$" + final_exup, expected_up, "Result Not Found");
                                break;
                            }
                            else
                            {
                                Assert.AreEqual("$" + xlRange.Cells[j, 4].Value2.ToString("N2"), expected_up, "Result Not Found");
                                break;
                            } 
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }
            }
        }

        //verify total price of a product present on cart
        public void verifyTotalPrice()
        {
            for (int i = 1; i < rowCount - 1; i++)
            {
                IWebElement site_prdname = driver.FindElement(By.XPath("//div/table/tbody/tr[" + i + "]/td[@class='text-left']/a"));
                string prdname = site_prdname.Text;
                IWebElement site_uprice = driver.FindElement(By.XPath("//div/table/tbody/tr[" + i + "]/td[6][@class='text-right']"));
                string expected_up = site_uprice.Text;
                for (int j = 2; j <= rowCount; j++)
                {
                    string excel_prdname = xlRange.Cells[j, 1].Value2.ToString();
                    double excel_unitprice = xlRange.Cells[j, 4].Value2;
                    double excel_quant = xlRange.Cells[j, 6].Value2;
                    if (prdname == excel_prdname)
                    {
                        if (prdname == "Apple Cinema 30\"")
                        {
                            double new_exup = (excel_unitprice + 63.6)*excel_quant;
                            string final_extp = new_exup.ToString("N2");
                            Assert.AreEqual("$" + final_extp, expected_up, "Result Not Found");
                            break;
                        }
                        if (prdname == "Product 8")
                        {
                            double new_exup = (excel_unitprice + 12)*excel_quant;
                            string final_extp = new_exup.ToString("N2");
                            Assert.AreEqual("$" + final_extp, expected_up, "Result Not Found");
                            break;
                        }
                        else
                        {
                            double new_exup = excel_unitprice * excel_quant;
                            string final_extp = new_exup.ToString("N2");
                            Assert.AreEqual("$" + final_extp, expected_up, "Result Not Found");
                            break;
                        }
                    }
                }
            }
        }

        public void ShoppingCart()
        {
            driver.Navigate().GoToUrl("http://tutorialsninja.com/demo/index.php?route=checkout/cart");
            Thread.Sleep(1000);
            //driver.FindElement(By.Id("cart-total")).Click();
            //driver.FindElement(By.XPath("//strong[text()=' View Cart']")).Click();
        }

        public void stopBrowser()
        {
            xlwb.Save();
            xlwb.Close();
            xlapp.Quit();
            driver.Quit();

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlws);

            //close and release
            Marshal.ReleaseComObject(xlwb);

            //quit and release
            Marshal.ReleaseComObject(xlapp);
            foreach (Process process in Process.GetProcessesByName("Excel"))
                process.Kill();
        }
    }
}
