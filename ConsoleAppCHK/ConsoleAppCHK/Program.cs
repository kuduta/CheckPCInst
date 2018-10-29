using System;
using System.IO;
using System.Management;
using System.Net;
using System.Net.Sockets;
using System.ServiceProcess;
using System.Text;
using Microsoft.Win32;

namespace ConsoleAppCHK
{
    class Program
    {

        public static string GetLocalIPAddress()
        {
            var host = Dns.GetHostEntry(Dns.GetHostName());
            foreach (var ip in host.AddressList)
            {
                if (ip.AddressFamily == AddressFamily.InterNetwork)
                {
                    return ip.ToString();
                }
            }
            throw new Exception("No network adapters with an IPv4 address in the system!");
        }

        static string ReadSubKeyValue(string subKey, string key)
        {
            string str = string.Empty;
            using (RegistryKey registryKey = Registry.LocalMachine.OpenSubKey(subKey))
            {
                if (registryKey != null)

                {
                    try
                    {
                        str = registryKey.GetValue(key).ToString();
                        registryKey.Close();
                    }
                    catch (Exception ex3)
            {
               // Console.WriteLine("Exception error with get regkey :" + ex1.ToString());
            }

                }
            }
            return str;
        }
        public static String GetWindowsServiceStatus(String SERVICENAME)
        {
            ServiceController sc = new ServiceController(SERVICENAME);

            switch (sc.Status)
            {
                case ServiceControllerStatus.Running:
                    return "Running";
                case ServiceControllerStatus.Stopped:
                    return "Stopped";
                case ServiceControllerStatus.Paused:
                    return "Paused";
                case ServiceControllerStatus.StopPending:
                    return "Stopping";
                case ServiceControllerStatus.StartPending:
                    return "Starting";
                default:
                    return "Status Changing";
            }
        }
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
            //string regtry_patch = "SOFTWARE\\Lumension\\LMAgent";
            string EMSSagentstatus = "";
            string AVstatus = "";
            
           

            ComName = Environment.MachineName.ToString();

            if (Environment.Is64BitOperatingSystem)
            {
                regtry_text = @"SOFTWARE\WOW6432Node\TrendMicro\PC-cillinNTCorp\CurrentVersion";
            }
            else
            {
                regtry_text = @"SOFTWARE\TrendMicro\PC-cillinNTCorp\CurrentVersion";
            }
            

            ipaddress = ReadSubKeyValue(regtry_text, "IP"); 
            NetID = ipaddress.Substring(0, ipaddress.LastIndexOf("."));
            MAC = ReadSubKeyValue(regtry_text, "MAC"); 
            GUID = ReadSubKeyValue(regtry_text, "GUID"); 
            InstallDate = ReadSubKeyValue(regtry_text, "InstDate"); 
            domain = ReadSubKeyValue(regtry_text, "Domain");
            NtVer = ReadSubKeyValue(regtry_text, "NtVer");
            server = ReadSubKeyValue(regtry_text, "Server");
            AVstatus = GetWindowsServiceStatus("ntrtscan");
            EMSSagentstatus = GetWindowsServiceStatus("EMSS Agent");


            try
            {
                RegistryKey masterKey = Registry.LocalMachine.OpenSubKey(regtry_text + "\\Misc.");
                if (masterKey != null)
                {
                    version = masterKey.GetValue("ProgramVer").ToString();
                    productname = masterKey.GetValue("ProductName").ToString();
                }
                masterKey.Close(); 
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

            Console.WriteLine("EMSS Agent status:  " + EMSSagentstatus);
            Console.WriteLine("AV status   :  " + AVstatus);
            
            
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
                    
                    w.WriteLine("<table style=\"width:60%\" border=\"0\">");
                    w.WriteLine("<tr><td><H5>Computer Name  : </td><td> " + ComName + "</td></tr>");
                    w.WriteLine("<tr><td><H4>Serial Number :  </td><td>" + SerialNumber + " </td ></tr> ");
                    w.WriteLine("<tr><td><H4>SID :  </td><td> " + SID + "</td></tr>");
                    w.WriteLine("<tr><td><H4>Manufacturer  :  </td><td> " + Manufacturer + "</td></tr>");
                    w.WriteLine("<tr><td><H4>Windows version  :  </td><td> " + NtVer + "</td></tr>");
                    w.WriteLine("<tr><td><H4>OS  :  </td><td> " + OperSys + "</td></tr>");
                    
                    w.WriteLine("<tr><td><H4>Group  :  </td><td> " + domain + "</td></tr>");
                    w.WriteLine("<tr><td><H4>IP Address  :  </td><td> " + ipaddress + "</td></tr>");
                    //w.WriteLine("<tr><td>Net ID </td><td> " + NetID + "</td></tr>");
                    w.WriteLine("<tr><td><H4>MAC Address  :  </td><td> " + MAC + "</td></tr>");
                    w.WriteLine("<tr><td><H4>Antivirus Product  :  </td><td> " + productname + "</td></tr>");
                    w.WriteLine("<tr><td><H4>Antivirus Version  :  </td><td> " + version + "</td></tr>");
                    w.WriteLine("<tr><td><H4>Antivirus GUID  :  </td><td> " + GUID + "</td></tr>");
                    w.WriteLine("<tr><td><H4>Antivirus Server  :  </td><td> " + server + "</td></tr>");
                    w.WriteLine("<tr><td><H4>Antivirus status  :  </td><td> " + AVstatus + "</td></tr>");
                    w.WriteLine("<tr><td><H4>EMSS Agent status  :  </td><td> " + EMSSagentstatus + "</td></tr>");
                    

                   
                    w.WriteLine(" <form action = \"http://122.155.190.111/input\" medthod =\"POST\">");
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
                   w.WriteLine("</table>");
                    w.WriteLine("<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp<button align=\"center\" type = \"submit\" formmethod = \"post\" > Send Data </ button >");
                    
                    
                    w.WriteLine("</form >");
                    
                }
            }
            System.Diagnostics.Process.Start("ComInfo.html");




            Console.WriteLine("Step Out Press any key to exit.");
            //Console.ReadKey();


        }
    }
}
