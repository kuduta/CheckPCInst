using System;
using System.IO;
using System.Management;
using System.Text;
using Microsoft.Win32;

namespace ConsoleAppCHK
{
    class Program
    {
        static void Main(string[] args)
        {
            string regtry_text = "";
            string SID = "";
            string ComName = "";
            string Manufacturer = "";
            string OperSys = "";
            string SerialNumber = "";
            
            string version = "";
            string productname = "";
            string ipaddress = "";
            string MAC = "";
            string GUID = "";
            string InstallDate = "";
            string domain = "";
            string NtVer = "";
            string server = "";
            string NetID = "";
           

            ComName = Environment.MachineName.ToString();

            if (Environment.Is64BitOperatingSystem)
            {
                regtry_text = "SOFTWARE\\WOW6432Node\\TrendMicro\\PC-cillinNTCorp\\CurrentVersion";
            }
            else
            {
                regtry_text = "SOFTWARE\\TrendMicro\\PC-cillinNTCorp\\CurrentVersion";
            }

            try
            {

                RegistryKey regkey = Registry.LocalMachine.OpenSubKey(regtry_text);


                if (regkey != null)
                {
                    ipaddress = regkey.GetValue("IP").ToString();
                    NetID = ipaddress.Substring(0, ipaddress.LastIndexOf("."));
                    MAC = regkey.GetValue("MAC").ToString();
                    GUID = regkey.GetValue("GUID").ToString();
                    InstallDate = regkey.GetValue("InstDate").ToString();
                    domain = regkey.GetValue("Domain").ToString();
                    NtVer = regkey.GetValue("NtVer").ToString();
                    server = regkey.GetValue("Server").ToString();
                }
                regkey.Close();
            }
            catch (Exception ex1)
            {
               // Console.WriteLine("Exception error with get regkey :" + ex1.ToString());
            }


            try
            {
                RegistryKey masterKey = Registry.LocalMachine.OpenSubKey(regtry_text + "\\Misc.");



                if (masterKey != null)
                {
                    version = masterKey.GetValue("ProgramVer").ToString();
                    productname = masterKey.GetValue("ProductName").ToString();

                }
                masterKey.Close();

                

                //Console.WriteLine("Step  In Press any key to exit.");
                // Console.ReadKey();
            }
            catch (Exception ex)
            {
                //Console.WriteLine("Exception error with create htm :" + ex.ToString());
            }





            ConnectionOptions options = new ConnectionOptions();
            options.Impersonation = System.Management.ImpersonationLevel.Impersonate;

            ManagementScope scope = new ManagementScope("\\\\.\\root\\cimv2", options);
            scope.Connect();

            // get information of machine
            ObjectQuery query = new ObjectQuery("SELECT * FROM Win32_OperatingSystem");
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, query);
            ManagementObjectCollection queryCollection = searcher.Get();
            foreach (ManagementObject m in queryCollection)
            {
                
                OperSys = m["Caption"].ToString().Trim();
            }


            ObjectQuery queryBIOS = new ObjectQuery("SELECT * FROM Win32_BIOS");
            ManagementObjectSearcher searcherBIOS = new ManagementObjectSearcher(scope, queryBIOS);
            ManagementObjectCollection queryCollectionBIOS = searcherBIOS.Get();
            foreach (ManagementObject m in queryCollectionBIOS)
            {
                Manufacturer = m["Manufacturer"].ToString().Trim();
                SerialNumber = m["SerialNumber"].ToString().Trim();
            }
                      
            //Get SID in machine 
            ObjectQuery queryAcc = new ObjectQuery("SELECT * FROM Win32_UserAccount");
            ManagementObjectSearcher searcherAcc = new ManagementObjectSearcher(scope, queryAcc);
            ManagementObjectCollection queryCollectionAcc = searcherAcc.Get();
            foreach (ManagementObject m in queryCollectionAcc)
            {
                SID = m["SID"].ToString().Trim().Substring(0, 40);
            }

            Console.WriteLine("Computer name  : {0}", ComName);
            Console.WriteLine("Operating System   : {0}", OperSys);
            Console.WriteLine("Manufacturer  : {0}", Manufacturer);
            Console.WriteLine("SerialNumber : {0}", SerialNumber);
            Console.WriteLine("SID  : {0}", SID);

            Console.WriteLine("IP Address is  :  " + ipaddress);
            Console.WriteLine("IP Cut is  :  " + NetID);
            Console.WriteLine("MAC Address is  :  " + MAC);
            Console.WriteLine("GUID is  :  " + GUID);
            Console.WriteLine("Domain is  :  " + domain);
            Console.WriteLine("Windows NT version is  :  " + NtVer);
            Console.WriteLine("Server is  :  " + server);
          
            Console.WriteLine("Product Name is  :  " + productname);
            Console.WriteLine("Version is :  " + version);
            Console.WriteLine("Install Date :  " + InstallDate);
            
            
            Console.WriteLine();
            Console.WriteLine("Get information finish");
           

            using (FileStream fs = new FileStream("ComInfo.html", FileMode.Create))
            {
                using (StreamWriter w = new StreamWriter(fs, Encoding.UTF8))
                {
                    //ipaddress, MAC, GUID, InstallDate, domain, NtVer, server
                    w.WriteLine("<H1>System Information</H1>");

                    //w.WriteLine("<br>Computer Name is  :  " + ComName);
                    //w.WriteLine("<br>Serial Number is  :  " + SerialNumber);
                    //w.WriteLine("<br>SID is  : " + SID);
                    //w.WriteLine("<br>Manufacturer  :  " + Manufacturer);
                    //w.WriteLine("<br>Windows NT version is  :  " + NtVer);
                    //w.WriteLine("<br>OS is  :  " + OperSys);
                    //w.WriteLine("<br>Group is  :  " + domain);
                    //w.WriteLine("<br>IP Address is  :  " + ipaddress);
                    //w.WriteLine("<br>Net ID  :  " + NetID);
                    //w.WriteLine("<br>MAC Address is  :  " + MAC);              
                    //w.WriteLine("<br>Product Name is  :  " + productname);
                    //w.WriteLine("<br>Version is :  " + version);
                    //w.WriteLine("<br>GUID is  :  " + GUID);
                    //w.WriteLine("<br>Server AV is  :  " + server);

                    w.WriteLine("<table style=\"width:70%\">");
                    w.WriteLine("<tr><td>Computer Name </td><td> " + ComName + "</td></tr>");
                    w.WriteLine("<tr><td>Serial Number </td><td>" + SerialNumber + " </td ></tr> ");
                    w.WriteLine("<tr><td>SID </td><td> " + SID + "</td></tr>");
                    w.WriteLine("<tr><td>Manufacturer</td><td> " + Manufacturer + "</td></tr>");
                    w.WriteLine("<tr><td>Windows version</td><td> " + NtVer + "</td></tr>");
                    w.WriteLine("<tr><td>OS</td><td> " + OperSys + "</td></tr>");
                    
                    w.WriteLine("<tr><td>Group</td><td> " + domain + "</td></tr>");
                    w.WriteLine("<tr><td>IP Address</td><td> " + ipaddress + "</td></tr>");
                    w.WriteLine("<tr><td>Net ID </td><td> " + NetID + "</td></tr>");
                    w.WriteLine("<tr><td>MAC Address </td><td> " + MAC + "</td></tr>");
                    w.WriteLine("<tr><td>Product Name </td><td> " + productname + "</td></tr>");
                    w.WriteLine("<tr><td>Version </td><td> " + version + "</td></tr>");
                    w.WriteLine("<tr><td>GUID </td><td> " + GUID + "</td></tr>");
                    w.WriteLine("<tr><td>Server AV </td><td> " + server + "</td></tr>");
                    w.WriteLine("</table>");
                    w.WriteLine(" <form action = \"http://112.155.190.111/input\" medthod =\"POST\"");
                    w.WriteLine("<input type=\"hidden\" name=\"ComName\" value=\""+ ComName+ "\">");
                    w.WriteLine("<input type=\"hidden\" name=\"SerialNumber\" value = \""+SerialNumber+"\">");
                    w.WriteLine("<input type=\"hidden\" name=\"SID\" value = \""+SID+"\">");
                    w.WriteLine("<input type=\"hidden\" name=\"Manufacturer\" value = \""+Manufacturer+"\">");
                    w.WriteLine("<input type=\"hidden\" name=\"NtVer\" value = \""+NtVer+"\">");
                    w.WriteLine("<input type=\"hidden\" name=\"OperSys\" value = \""+OperSys+"\">");
                    w.WriteLine("<input type=\"hidden\" name=\"domain\" value = \""+domain+"\">");
                    w.WriteLine("<input type=\"hidden\" name=\"ipaddress\" value = \""+ipaddress+"\">");
                    w.WriteLine("<input type=\"hidden\" name=\"NetID\" value = \""+NetID+"\">");
                    w.WriteLine("<input type=\"hidden\" name=\"MAC\" value = \""+MAC+"\">");
                    w.WriteLine("<input type=\"hidden\" name=\"productname\" value = \""+productname+"\">");
                    w.WriteLine("<input type=\"hidden\" name=\"version\" value = \""+version+"\">");
                    w.WriteLine("<input type=\"hidden\" name=\"GUID\" value = \""+GUID+"\">");
                    w.WriteLine("<input type=\"hidden\" name=\"server\" value = \""+server+"\">");
                    
                    w.WriteLine("<button type = \"submit\" formmethod = \"post\" > Send Data </ button >");
                    w.WriteLine("</form >");
                    
                }
            }
            System.Diagnostics.Process.Start("ComInfo.html");



            Console.WriteLine("Step Out Press any key to exit.");
            Console.ReadKey();


        }
    }
}
