using System;
using System.Xml;
using Microsoft.BizTalk.ExplorerOM;
using Microsoft.Win32;

namespace Microsys.EAI.Framework.PasswordManager
{
    class Program
    {

        private static string _commandName = string.Empty;
        private static string _applicationName = string.Empty;
        private static string _portType = string.Empty;
        private static string _portName = string.Empty;
        private static string _userName = string.Empty;
        private static string _password = string.Empty;

        static void Main(string[] args)
        {

            ReadParameters(args);

            if (string.IsNullOrEmpty(_commandName))
            {
                PrintHelp();
                return;
            }

            switch (_commandName)
            {

                case "-list":

                    if (string.IsNullOrEmpty(_applicationName))
                    {
                        PrintHelp();
                        return;
                    }

                    List(_applicationName);

                    break;

                case "-get":

                    if (string.IsNullOrEmpty(_portType) || string.IsNullOrEmpty(_portName))
                    {
                        PrintHelp();
                        return;
                    }

                    if (_portType == "-receive")
                    {
                        GetReceiveLocation(_portName);
                    }
                    else
                    {
                        GetSendPort(_portName);
                    }

                    break;

                case "-set":

                    if (string.IsNullOrEmpty(_portType) || string.IsNullOrEmpty(_portName) || string.IsNullOrEmpty(_userName) || string.IsNullOrEmpty(_password))
                    {
                        PrintHelp();
                        return;
                    }

                    if (_portType == "-receive")
                    {
                        SetReceiveLocation(_portName, _userName, _password);
                    }
                    else
                    {
                        SetSendPort(_portName, _userName, _password);
                    }

                    break;

            }


        }

        static void ReadParameters(string[] args)
        {

            foreach (string param in args)
            {

                var arg = param.ToLower();

                if (arg == "-list" || arg == "-get" || arg == "-set")
                {
                    _commandName = arg;
                }

                if (arg.StartsWith("-application:"))
                {
                    _applicationName = param.Split(":".ToCharArray())[1];
                }

                if (arg == "-send" || arg == "-receive")
                {
                    _portType = arg;
                }

                if (arg.StartsWith("-name:"))
                {
                    _portName = param.Split(":".ToCharArray())[1];
                }

                if (arg.StartsWith("-user:"))
                {
                    _userName = param.Split(":".ToCharArray())[1];
                }

                if (arg.StartsWith("-password:"))
                {
                    _password = param.Split(":".ToCharArray())[1];
                }

            }

        }

        static void PrintHelp()
        {
            Console.WriteLine("Missing parameters\r\n");
            Console.WriteLine("Parameters:\r\n");
            Console.WriteLine("-list -application:[application name]\r\n");
            Console.WriteLine("-get -receive -name:[receive location name]\r\n");
            Console.WriteLine("-get -send -name:[send port name]\r\n");
            Console.WriteLine("-set -receive -name:[receive location name] -user:[username] -password:[password]\r\n");
            Console.WriteLine("-set -send -name:[send port name] -user:[username] -password:[password]\r\n");
        }

        private static string GetConnectionString()
        {

            string connectionString = string.Empty;

            using (RegistryKey bizTalkAdminKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\BizTalk Server\3.0\Administration"))
            {
                if (bizTalkAdminKey != null)
                {
                    string server = (string)bizTalkAdminKey.GetValue("MgmtDBServer", ".");
                    string database = (string)bizTalkAdminKey.GetValue("MgmtDBName", "BizTalkMgmtDb");

                    connectionString = string.Format("Data Source={0};Initial Catalog={1};Integrated Security=SSPI;", server, database);
                }
            }

            return connectionString;

        }

        private static string GetTransportTypeData(string transportTypeData)
        {

            XmlDocument transportTypeDataXml = new XmlDocument();
            transportTypeDataXml.LoadXml(transportTypeData);

            return transportTypeDataXml.InnerXml;
        }

        private static void List(string applicationName)
        {

            BtsCatalogExplorer catalog = new BtsCatalogExplorer();
            catalog.ConnectionString = GetConnectionString();

            Application application = catalog.Applications[applicationName];

            Console.WriteLine("Receive Locations:");

            foreach (ReceivePort receivePort in application.ReceivePorts)
            {
                foreach (ReceiveLocation receiveLocation in receivePort.ReceiveLocations)
                {
                    Console.WriteLine("-----------------------------------------------------------");
                    Console.WriteLine("{0}.{1} ({2})", application.Name, receiveLocation.Name, receiveLocation.TransportType.Name);
                    Console.WriteLine(GetTransportTypeData(receiveLocation.TransportTypeData));
                }

            }

            Console.WriteLine("Send Ports:");

            foreach (SendPort sendPort in application.SendPorts)
            {
                if (sendPort.PrimaryTransport != null)
                {
                    Console.WriteLine("-----------------------------------------------------------");
                    Console.WriteLine("{0}.{1} ({2})", application.Name, sendPort.Name, sendPort.PrimaryTransport.TransportType.Name);
                    Console.WriteLine(GetTransportTypeData(sendPort.PrimaryTransport.TransportTypeData));
                }
                if (sendPort.SecondaryTransport != null && sendPort.SecondaryTransport.TransportType != null)
                {
                    Console.WriteLine("-----------------------------------------------------------");
                    Console.WriteLine("{0}.{1} ({2})", application.Name, sendPort.Name, sendPort.SecondaryTransport.TransportType.Name);
                    Console.WriteLine(GetTransportTypeData(sendPort.SecondaryTransport.TransportTypeData));
                }
            }


        }

        private static void GetReceiveLocation(string receiveLocationName)
        {

            BtsCatalogExplorer catalog = new BtsCatalogExplorer();
            catalog.ConnectionString = GetConnectionString();

            foreach (ReceivePort receivePort in catalog.ReceivePorts)
            {
                foreach (ReceiveLocation receiveLocation in receivePort.ReceiveLocations)
                {
                    if (receiveLocation.Name.ToLower() == receiveLocationName.ToLower())
                    {
                        Console.WriteLine("{0} ({1})", receiveLocation.Name, receiveLocation.TransportType.Name);
                        Console.WriteLine(GetTransportTypeData(receiveLocation.TransportTypeData));
                    }
                }
            }

        }

        private static void GetSendPort(string sendPortName)
        {

            BtsCatalogExplorer catalog = new BtsCatalogExplorer();
            catalog.ConnectionString = GetConnectionString();

            foreach (SendPort sendPort in catalog.SendPorts)
            {

                if (sendPort.Name.ToLower() == sendPortName.ToLower())
                {

                    if (sendPort.PrimaryTransport != null)
                    {
                        Console.WriteLine("{0} ({1})", sendPort.Name, sendPort.PrimaryTransport.TransportType.Name);
                        Console.WriteLine(GetTransportTypeData(sendPort.PrimaryTransport.TransportTypeData));
                    }
                    if (sendPort.SecondaryTransport != null && sendPort.SecondaryTransport.TransportType != null)
                    {
                        Console.WriteLine("{0} ({1})", sendPort.Name, sendPort.SecondaryTransport.TransportType.Name);
                        Console.WriteLine(GetTransportTypeData(sendPort.SecondaryTransport.TransportTypeData));
                    }

                }
            }


        }

        private static void SetReceiveLocation(string receiveLocationName, string userName, string password)
        {

            bool found = false;

            try
            {

                BtsCatalogExplorer catalog = new BtsCatalogExplorer();
                catalog.ConnectionString = GetConnectionString();

                foreach (ReceivePort receivePort in catalog.ReceivePorts)
                {
                    foreach (ReceiveLocation receiveLocation in receivePort.ReceiveLocations)
                    {
                        if (receiveLocation.Name.ToLower() == receiveLocationName.ToLower())
                        {

                            found = true;

                            XmlDocument transportTypeData = new XmlDocument();
                            transportTypeData.LoadXml(receiveLocation.TransportTypeData);

                            LogHelper.Write(LogHelper.EntryType.Information, "Receive Location '{0}' ({1})", receiveLocation.Name, receiveLocation.TransportType.Name);
                            LogHelper.Write(LogHelper.EntryType.Information, "Original configuration:");
                            LogHelper.Write(LogHelper.EntryType.Information, transportTypeData.InnerXml);

                            switch (receiveLocation.TransportType.Name)
                            {
                                case "FILE":

                                    SetTransportDataForFile(transportTypeData, userName, password);

                                    receiveLocation.TransportTypeData = transportTypeData.InnerXml;
                                    catalog.SaveChanges();

                                    break;

                                case "FTP":

                                    SetTransportDataForFtp(transportTypeData, userName, password);

                                    receiveLocation.TransportTypeData = transportTypeData.InnerXml;
                                    catalog.SaveChanges();

                                    break;

                                case "SFTP":

                                    SetTransportDataForSftp(transportTypeData, userName, password);

                                    receiveLocation.TransportTypeData = transportTypeData.InnerXml;
                                    catalog.SaveChanges();

                                    break;

                                default:

                                    if (receiveLocation.TransportType.Name.StartsWith("WCF"))
                                    {
                                        SetTransportDataForWcf(transportTypeData, userName, password);

                                        receiveLocation.TransportTypeData = transportTypeData.InnerXml;
                                        catalog.SaveChanges();
                                    }
                                    else
                                    {
                                        throw (new Exception(string.Format("Transport {0} not managed.", receiveLocation.TransportType.Name)));
                                    }

                                    break;

                            }

                            LogHelper.Write(LogHelper.EntryType.Information, "New configuration:");
                            LogHelper.Write(LogHelper.EntryType.Information, transportTypeData.InnerXml);

                            Console.WriteLine(receiveLocationName + " -> OK");

                        }
                    }
                }

                if (!found)
                {
                    throw (new Exception(string.Format("Receive Location {0} not found.", receiveLocationName)));
                }

            }
            catch (Exception exc)
            {

                LogHelper.Write(LogHelper.EntryType.Error, exc.Message);

                ConsoleColor currentColor = Console.ForegroundColor;
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(receiveLocationName + " -> KO");
                Console.ForegroundColor = currentColor;
            }

        }

        private static void SetSendPort(string sendPortName, string userName, string password)
        {

            bool found = false;

            try
            {

                BtsCatalogExplorer catalog = new BtsCatalogExplorer();
                catalog.ConnectionString = GetConnectionString();

                foreach (SendPort sendPort in catalog.SendPorts)
                {
                    if (sendPort.Name.ToLower() == sendPortName.ToLower())
                    {

                        found = true;

                        XmlDocument transportTypeData = new XmlDocument();
                        transportTypeData.LoadXml(sendPort.PrimaryTransport.TransportTypeData);

                        LogHelper.Write(LogHelper.EntryType.Information, "Send Port '{0}' ({1}):", sendPort.Name, sendPort.PrimaryTransport.TransportType.Name);
                        LogHelper.Write(LogHelper.EntryType.Information, "Original configuration:");
                        LogHelper.Write(LogHelper.EntryType.Information, transportTypeData.InnerXml);

                        switch (sendPort.PrimaryTransport.TransportType.Name)
                        {

                            case "FILE":

                                SetTransportDataForFile(transportTypeData, userName, password);

                                sendPort.PrimaryTransport.TransportTypeData = transportTypeData.InnerXml;
                                catalog.SaveChanges();

                                break;

                            case "FTP":

                                SetTransportDataForFtp(transportTypeData, userName, password);

                                sendPort.PrimaryTransport.TransportTypeData = transportTypeData.InnerXml;
                                catalog.SaveChanges();

                                break;

                            case "SFTP":

                                SetTransportDataForSftp(transportTypeData, userName, password);

                                sendPort.PrimaryTransport.TransportTypeData = transportTypeData.InnerXml;
                                catalog.SaveChanges();

                                break;

                            case "POP3":

                                SetTransportDataForPop3(transportTypeData, userName, password);

                                sendPort.PrimaryTransport.TransportTypeData = transportTypeData.InnerXml;
                                catalog.SaveChanges();

                                break;

                            default:

                                if (sendPort.PrimaryTransport.TransportType.Name.StartsWith("WCF"))
                                {
                                    SetTransportDataForWcf(transportTypeData, userName, password);

                                    sendPort.PrimaryTransport.TransportTypeData = transportTypeData.InnerXml;
                                    catalog.SaveChanges();
                                }
                                else
                                {
                                    throw (new Exception(string.Format("Transport {0} not managed.", sendPort.PrimaryTransport.TransportType.Name)));
                                }

                                break;


                        }

                        LogHelper.Write(LogHelper.EntryType.Information, "New configuration:");
                        LogHelper.Write(LogHelper.EntryType.Information, transportTypeData.InnerXml);

                        Console.WriteLine(sendPortName + " -> OK");

                    }

                }

                if (!found)
                {
                    throw (new Exception(string.Format("Send Port {0} not found.", sendPortName)));
                }

            }
            catch (Exception exc)
            {
                LogHelper.Write(LogHelper.EntryType.Error, exc.Message);

                ConsoleColor currentColor = Console.ForegroundColor;
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(sendPortName + " -> KO");
                Console.ForegroundColor = currentColor;
            }

        }

        private static void SetTransportDataForFile(XmlDocument originalTransportData, string userName, string password)
        {

            XmlNode userNameNode;
            XmlAttribute userNameAttribute;
            XmlNode passwordNode;
            XmlAttribute passwordAttribute;

            if (originalTransportData.SelectSingleNode(@"/CustomProps/Username") == null)
            {
                userNameNode = originalTransportData.SelectSingleNode(@"/CustomProps").AppendChild(originalTransportData.CreateElement("Username"));
                userNameAttribute = userNameNode.Attributes.Append(originalTransportData.CreateAttribute("vt"));
            }
            else
            {
                userNameNode = originalTransportData.SelectSingleNode(@"/CustomProps/Username");
                userNameAttribute = userNameNode.Attributes["vt"];

            }

            userNameNode.InnerXml = userName;
            userNameAttribute.Value = "8";

            if (originalTransportData.SelectSingleNode(@"/CustomProps/Password") == null)
            {
                passwordNode = originalTransportData.SelectSingleNode(@"/CustomProps").AppendChild(originalTransportData.CreateElement("Password"));
                passwordAttribute = passwordNode.Attributes.Append(originalTransportData.CreateAttribute("vt"));
            }
            else
            {
                passwordNode = originalTransportData.SelectSingleNode(@"/CustomProps/Password");
                passwordAttribute = passwordNode.Attributes["vt"];
            }

            passwordNode.InnerXml = password;
            passwordAttribute.Value = "8";

        }

        private static void SetTransportDataForWcf(XmlDocument originalTransportData, string userName, string password)
        {

            XmlNode userNameNode;
            XmlAttribute userNameAttribute;
            XmlNode passwordNode;
            XmlAttribute passwordAttribute;

            if (originalTransportData.SelectSingleNode(@"/CustomProps/UserName") == null)
            {
                userNameNode = originalTransportData.SelectSingleNode(@"/CustomProps").AppendChild(originalTransportData.CreateElement("UserName"));
                userNameAttribute = userNameNode.Attributes.Append(originalTransportData.CreateAttribute("vt"));
            }
            else
            {
                userNameNode = originalTransportData.SelectSingleNode(@"/CustomProps/UserName");
                userNameAttribute = userNameNode.Attributes["vt"];

            }

            userNameNode.InnerXml = userName;
            userNameAttribute.Value = "8";

            if (originalTransportData.SelectSingleNode(@"/CustomProps/Password") == null)
            {
                passwordNode = originalTransportData.SelectSingleNode(@"/CustomProps").AppendChild(originalTransportData.CreateElement("Password"));
                passwordAttribute = passwordNode.Attributes.Append(originalTransportData.CreateAttribute("vt"));
            }
            else
            {
                passwordNode = originalTransportData.SelectSingleNode(@"/CustomProps/Password");
                passwordAttribute = passwordNode.Attributes["vt"];
            }

            passwordNode.InnerXml = password;
            passwordAttribute.Value = "8";

        }

        private static void SetTransportDataForFtp(XmlDocument originalTransportData, string userName, string password)
        {

            XmlNode adapterConfigNode = originalTransportData.SelectSingleNode(@"/CustomProps/AdapterConfig");
            string adapterConfigString = adapterConfigNode.InnerXml;

            adapterConfigString = adapterConfigString.Replace("&lt;", "<");
            adapterConfigString = adapterConfigString.Replace("&gt;", ">");

            XmlDocument adapterConfigDocument = new XmlDocument();
            adapterConfigDocument.LoadXml(adapterConfigString);

            LogHelper.Write(LogHelper.EntryType.Information, "Adapter config original configuration:");
            LogHelper.Write(LogHelper.EntryType.Information, adapterConfigDocument.InnerXml);

            XmlNode userNameNode;

            if (adapterConfigDocument.SelectSingleNode(@"/Config/userName") == null)
            {
                userNameNode = adapterConfigDocument.SelectSingleNode(@"/Config").AppendChild(adapterConfigDocument.CreateElement("userName"));
            }
            else
            {
                userNameNode = adapterConfigDocument.SelectSingleNode(@"/Config/userName");

            }

            userNameNode.InnerXml = userName;

            XmlNode passwordNode;

            if (adapterConfigDocument.SelectSingleNode(@"/Config/password") == null)
            {
                passwordNode = adapterConfigDocument.SelectSingleNode(@"/Config").AppendChild(adapterConfigDocument.CreateElement("password"));
            }
            else
            {
                passwordNode = adapterConfigDocument.SelectSingleNode(@"/Config/password");

            }

            passwordNode.InnerXml = password;

            LogHelper.Write(LogHelper.EntryType.Information, "Adapter config new configuration:");
            LogHelper.Write(LogHelper.EntryType.Information, adapterConfigDocument.InnerXml);

            adapterConfigString = adapterConfigDocument.InnerXml;

            adapterConfigString = adapterConfigString.Replace("<", "&lt;");
            adapterConfigString = adapterConfigString.Replace(">", "&gt;");

            adapterConfigNode.InnerXml = adapterConfigString;
        }

        private static void SetTransportDataForSftp(XmlDocument originalTransportData, string userName, string password)
        {

            XmlNode userNameNode;
            XmlAttribute userNameAttribute;
            XmlNode passwordNode;
            XmlAttribute passwordAttribute;

            if (originalTransportData.SelectSingleNode(@"/CustomProps/UserName") == null)
            {
                userNameNode = originalTransportData.SelectSingleNode(@"/CustomProps").AppendChild(originalTransportData.CreateElement("UserName"));
                userNameAttribute = userNameNode.Attributes.Append(originalTransportData.CreateAttribute("vt"));
            }
            else
            {
                userNameNode = originalTransportData.SelectSingleNode(@"/CustomProps/UserName");
                userNameAttribute = userNameNode.Attributes["vt"];
            }

            userNameNode.InnerXml = userName;
            userNameAttribute.Value = "8";

            if (originalTransportData.SelectSingleNode(@"/CustomProps/Password") == null)
            {
                passwordNode = originalTransportData.SelectSingleNode(@"/CustomProps").AppendChild(originalTransportData.CreateElement("Password"));
                passwordAttribute = passwordNode.Attributes.Append(originalTransportData.CreateAttribute("vt"));
            }
            else
            {
                passwordNode = originalTransportData.SelectSingleNode(@"/CustomProps/Password");
                passwordAttribute = passwordNode.Attributes["vt"];
            }

            passwordNode.InnerXml = password;
            passwordAttribute.Value = "8";
        }

        private static void SetTransportDataForPop3(XmlDocument originalTransportData, string userName, string password)
        {

            XmlNode adapterConfigNode = originalTransportData.SelectSingleNode(@"/CustomProps/AdapterConfig");
            string adapterConfigString = adapterConfigNode.InnerXml;

            adapterConfigString = adapterConfigString.Replace("&lt;", "<");
            adapterConfigString = adapterConfigString.Replace("&gt;", ">");

            XmlDocument adapterConfigDocument = new XmlDocument();
            adapterConfigDocument.LoadXml(adapterConfigString);

            LogHelper.Write(LogHelper.EntryType.Information, "Adapter config original configuration:");
            LogHelper.Write(LogHelper.EntryType.Information, adapterConfigDocument.InnerXml);

            XmlNode userNameNode;


            if (adapterConfigDocument.SelectSingleNode(@"/Config/userName") == null)
            {
                userNameNode = adapterConfigDocument.SelectSingleNode(@"/Config").AppendChild(adapterConfigDocument.CreateElement("userName"));
            }
            else
            {
                userNameNode = adapterConfigDocument.SelectSingleNode(@"/Config/userName");

            }

            userNameNode.InnerXml = userName;

            XmlNode passwordNode;
            XmlAttribute passwordAttribute;

            if (adapterConfigDocument.SelectSingleNode(@"/Config/password") == null)
            {
                passwordNode = adapterConfigDocument.SelectSingleNode(@"/Config").AppendChild(adapterConfigDocument.CreateElement("password"));
                passwordAttribute = passwordNode.Attributes.Append(adapterConfigDocument.CreateAttribute("vt"));
            }
            else
            {
                passwordNode = adapterConfigDocument.SelectSingleNode(@"/Config/password");
                passwordAttribute = passwordNode.Attributes["vt"];
            }

            passwordNode.InnerXml = password;
            passwordAttribute.Value = "8";

            LogHelper.Write(LogHelper.EntryType.Information, "Adapter config new configuration:");
            LogHelper.Write(LogHelper.EntryType.Information, adapterConfigDocument.InnerXml);

            adapterConfigString = adapterConfigDocument.InnerXml;

            adapterConfigString = adapterConfigString.Replace("<", "&lt;");
            adapterConfigString = adapterConfigString.Replace(">", "&gt;");

            adapterConfigNode.InnerXml = adapterConfigString;
        }
    }
}
