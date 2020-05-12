
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Windows.Forms;


using Microsoft.SharePoint.Client;
using Microsoft.Online.SharePoint.TenantAdministration;

using OfficeDevPnP.Core.Entities;
using CsvHelper;
using AngleSharp.Dom.Css;

namespace GetSiteCollectionsM365
{
    //Defined for use of CSV Helper
    public class M365SiteCollection
    {
        public string Name { get; set; }
        public string Url { get; set; }
        public DateTime LastModified { get; set; }
    }

    class GetSiteCollections
    {
        [STAThread]
        static void Main(string[] args)
        {
            //https://suniradev-admin.sharepoint.com/ - Test SharePoint Tenant URL
            
            Console.WriteLine("Attempting to Log into M365! ");
            LoginUsingM365();
        }
        private static void LoginUsingM365()
        {
            string siteURL;
            bool isValidTenantURL;

            //Retrieve Admin Site URL from User via Console
            Console.WriteLine("Please enter the M365 Admin Site URL.");

            //Verify it's a valid Tenant Site URL
            do{
                siteURL = Console.ReadLine();
                isValidTenantURL = ValidateTenantURL(siteURL);
            } while (!isValidTenantURL);

            //Authenticate using M365
            var authenticationManager = new OfficeDevPnP.Core.AuthenticationManager();
            ClientContext context = authenticationManager.GetWebLoginClientContext(siteURL, null);
            
            GetAllTenantWebs(context);
            
        }

        private static void GetAllTenantWebs(ClientContext context)
        {
            /*Get Site Collections under Tenant, including OneDrive Sites*/
            Tenant tenant = new Tenant(context);
            IList<SiteEntity> siteCols = tenant.GetSiteCollections(startIndex: 0, includeDetail: true, includeOD4BSites: true);

            var records = new List<M365SiteCollection>();

            foreach (var siteCol in siteCols)
            {
                string siteTitle = siteCol.Title;
                string siteUrl = siteCol.Url.ToString();
                DateTime lastMod = siteCol.LastContentModifiedDate;

                siteTitle = siteUrl.Contains("-my.sharepoint.com") ? "Onedrive for Business" :
                    siteUrl.Contains(".sharepoint.com/search") ? "Search" :
                            siteUrl.Length < 1 ? "No Title" :
                            siteTitle; 
                
                Console.WriteLine(siteTitle + " : " + siteUrl + " Modified: " + lastMod.ToString());
                records.Add(new M365SiteCollection { 
                    Name = siteTitle, 
                    Url = siteCol.Url.ToString(), 
                    LastModified = lastMod 
                });
            }

            SaveToCSV(records);
        }

        private static void SaveToCSV(List<M365SiteCollection> records)
        {
            string fileName = "site-collections";
            string filePath = "/";
            Console.WriteLine("Please Select the Folder to save the CSV file. Press Enter to Continue.");
            while (!(Console.KeyAvailable && Console.ReadKey(true).Key == ConsoleKey.Enter))
            {
                //Just Hang Out until I get a keypress.
            }

            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                filePath = fbd.SelectedPath;
            }

            do
            {
                Console.WriteLine("Please give the .csv file a name, without the extension.");
                fileName = Console.ReadLine();
            } while (fileName.Length < 1);

            string fileRoute = filePath + "/" + fileName + ".csv";
            Console.WriteLine("Saving to :" + fileRoute);

            using (var writer = new StreamWriter(fileRoute))
            {
                using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
                {
                    csv.WriteRecords(records);
                }
            }
        }

        private static bool ValidateTenantURL(string siteURL)
        {
            bool isValidTenantURL = true;

            if ((siteURL.Length < 1))
            {
                Console.WriteLine("Not a valid URL.");
                isValidTenantURL = false;
            }

            if (!siteURL.Contains("https://"))
            {
                Console.WriteLine("Please enter a URL that starts with https://");
                isValidTenantURL = false;
            }

            if (!siteURL.Contains("-admin.sharepoint.com"))
            {
                Console.WriteLine("The site should end with '-admin.sharepoint.com/'");
                isValidTenantURL = false;
            }

            return isValidTenantURL;
        }

    }
}
