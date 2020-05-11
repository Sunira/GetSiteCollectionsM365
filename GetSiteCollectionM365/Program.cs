
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

using Microsoft.SharePoint.Client;
using Microsoft.Online.SharePoint.TenantAdministration;

using OfficeDevPnP.Core.Entities;


namespace GetSiteCollectionsM365
{
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
            IList<SiteEntity> siteCols = tenant.GetSiteCollections(startIndex: 0,
                                                     includeDetail: true,
                                                     includeOD4BSites: true);

            foreach(var siteCol in siteCols)
            {
                string title;

                title = siteCol.Url.ToString().Contains("-my.sharepoint.com/") ? "Onedrive Site" :
                        siteCol.Url.ToString().Length < 1 ? "No Title" : 
                        siteCol.Title; 
                
                Console.WriteLine(title + " : " + siteCol.Url.ToString() + " Modified: " + siteCol.LastContentModifiedDate);
            }

            SaveToCSV();
        }

        private static void SaveToCSV()
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                Console.WriteLine(fbd.SelectedPath); // full path
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
