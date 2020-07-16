using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportAppServer.DataDefModel;
using System;
using System.Data.SqlClient;
using System.IO;
using System.Linq;

namespace CrystalReportsOLEDB
{
    /// <summary>
    /// This program is based on the SAP' Knowledge Database.
    /// https://apps.support.sap.com/sap/support/knowledge/public/en/1535347
    /// </summary>
    class Program
    {
        private static readonly ApplicationConnectionInfo ConnectionInfo = new ApplicationConnectionInfo();
        static string path = "";

        static void Main(string[] args)
        {
            // Test if input arguments were supplied.
            if (args == null || args.Length == 0)
            {
                Console.WriteLine("No arguments found. Starting interactive mode...");
                Console.Write("Crystal Reports Path (recusive search is enabled): ");
                path = Console.ReadLine();
                Console.Write("Database Server: ");
                ConnectionInfo.Server = Console.ReadLine();
                Console.Write("Database Name: ");
                ConnectionInfo.Database = Console.ReadLine();
                Console.Write("Database Username: ");
                ConnectionInfo.Username = Console.ReadLine();
                Console.Write("Database Password: ");
                ConnectionInfo.Password = Console.ReadLine();
            }
            else
            {
                ProcessArguments(args);
            }

            Console.Clear();
            Console.WriteLine("Testing connection...");

            if (!TestConnection())
            {
                Console.WriteLine("Couldn't connect to SQL Server. Make sure that the provided credentials are valid and the firewall rules allow you to reach this server.");
                Console.ReadLine();

                return;
            }

            Console.Clear();
            TryToUpdateReports();

            Console.WriteLine("Press enter to close...");
            Console.ReadLine();

            return;
        }

        private static void ProcessArguments(string[] args)
        {
            for (int i = 0; i < args.Length; i++)
            {
                switch (args[i])
                {
                    case "-server":
                        ConnectionInfo.Server = args[i];
                        break;

                    case "-path":
                        path = args[i];
                        break;

                    case "-username":
                        ConnectionInfo.Username = args[i];
                        break;

                    case "-password":
                        ConnectionInfo.Password = args[i];
                        break;


                    case "-database":
                        ConnectionInfo.Database = args[i];
                        break;


                    case "-?":
                    case "-h":
                    case "-help":
                        Console.WriteLine("-server: Server name");
                        Console.WriteLine("-database: database name");
                        Console.WriteLine("-username: username");
                        Console.WriteLine("-password: password of the user");
                        Console.WriteLine("-path: root path of crystal reports files");
                        Console.ReadLine();
                        break;

                    default:
                        throw new ArgumentException(args[i]);
                }
            }
        }

        private static bool TestConnection()
        {
            string connectionString = $"Data Source={ConnectionInfo.Server};Initial Catalog={ConnectionInfo.Database};User id={ConnectionInfo.Username};Password={ConnectionInfo.Password};";
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    Console.WriteLine("Connection successful!");
                    return true;
                }
                catch (SqlException)
                {
                    return false;
                }
            }
        }

        private static void TryToUpdateReports()
        {

            string[] filePaths = Directory.GetFiles(path, "*.rpt", SearchOption.AllDirectories);
            int totalFiles = filePaths.Count();
            int failedFiles = 0;

            for (int i = 0; i < filePaths.Length; i++)
            {
                try
                {
                    Console.WriteLine($"Processing file in path {filePaths[i]}.");
                    UpdateReportConnection(filePaths[i]);
                }
                catch (Exception e)
                {
                    Console.WriteLine($"File in path '{filePaths[i]}' failed with the error: '{e.Message}'.");
                    Console.WriteLine($" Skipping to the next file...");
                    failedFiles++;
                }
            }

            Console.WriteLine($"Successfully processed {totalFiles - failedFiles}/{totalFiles}.");
        }

        private static void UpdateReportConnection(string reportPath)
        {
            var doc = new ReportDocument();
            doc.Load(reportPath);

            ReplaceReportConnection(doc);

            doc.SaveAs(reportPath, false);
            doc.Close();

            Console.WriteLine($"'{new FileInfo(reportPath).Name}' updated to MSOLEDBSQL.");
        }

        private static void ReplaceReportConnection(ReportDocument doc)
        {
            PropertyBag connectionAttributes = new PropertyBag();
            connectionAttributes.Add("Auto Translate", "-1");
            connectionAttributes.Add("Connect Timeout", "15");
            connectionAttributes.Add("Data Source", ConnectionInfo.Server);
            connectionAttributes.Add("General Timeout", "0");
            connectionAttributes.Add("Initial Catalog", ConnectionInfo.Database);
            connectionAttributes.Add("Integrated Security", false);
            connectionAttributes.Add("Locale Identifier", "1033");
            connectionAttributes.Add("OLE DB Services", "-5");
            connectionAttributes.Add("Provider", "MSOLEDBSQL");
            connectionAttributes.Add("Tag with column collation when possible", "0");
            connectionAttributes.Add("Use DSN Default Properties", false);
            connectionAttributes.Add("Use Encryption for Data", "0");

            PropertyBag attributes = new PropertyBag();
            attributes.Add("Database DLL", "crdb_ado.dll");
            attributes.Add("QE_DatabaseName", ConnectionInfo.Database);
            attributes.Add("QE_DatabaseType", "OLE DB (ADO)");
            attributes.Add("QE_LogonProperties", connectionAttributes);
            attributes.Add("QE_ServerDescription", ConnectionInfo.Server);
            attributes.Add("QESQLDB", true);
            attributes.Add("SSO Enabled", false);

            ConnectionInfo ci = new ConnectionInfo
            {
                Attributes = attributes,
                Kind = CrConnectionInfoKindEnum.crConnectionInfoKindCRQE,
                UserName = ConnectionInfo.Username,
                Password = ConnectionInfo.Password
            };

            foreach (CrystalDecisions.ReportAppServer.DataDefModel.Table table in doc.ReportClientDocument.DatabaseController.Database.Tables)
            {
                Procedure newTable = new Procedure
                {
                    ConnectionInfo = ci,
                    Name = table.Name,
                    Alias = table.Alias,
                    QualifiedName = ConnectionInfo.Database + ".dbo." + table.Name
                };

                doc.ReportClientDocument.DatabaseController.SetTableLocation(table, newTable);
            }

            foreach (ReportDocument subreport in doc.Subreports)
            {
                foreach (CrystalDecisions.ReportAppServer.DataDefModel.Table table in doc.ReportClientDocument.SubreportController.GetSubreportDatabase(subreport.Name).Tables)
                {
                    Procedure newTable = new Procedure
                    {
                        ConnectionInfo = ci,
                        Name = table.Name,
                        Alias = table.Alias,
                        QualifiedName = ConnectionInfo.Database + ".dbo." + table.Name
                    };
                    doc.ReportClientDocument.SubreportController.SetTableLocation(subreport.Name, table, newTable);
                }
            }
        }
    }
}
