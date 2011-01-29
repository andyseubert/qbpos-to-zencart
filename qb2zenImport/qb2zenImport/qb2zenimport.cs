using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Data.Common;
using MySql.Data.MySqlClient;
using System.Configuration;


namespace qb2zenImport
{
    class ImportToZencart
    {
        static void Main(string[] args)
        {
            string DTS = DateTime.Now.ToString("MMddyyhhmmss");
            string logPath = ConfigurationSettings.AppSettings["LogPath"];
            string notOnlineLog = logPath + DTS + "NotOnline.csv";
            string onlineLog = logPath + DTS + "Online.csv";
            string inStoreName;
            string inStoreNumber;
            string inStoreQTY;
            string zcName, zcNumber, zcQTY;

            // create connection to excel spreadsheet
            string thefile =  ConfigurationSettings.AppSettings["ItemsFile"];
            string filename = System.IO.Path.GetFileName(thefile);


            // check that this file exists before continuing
            if (!(File.Exists(thefile))) { Environment.Exit(0); }


            //open logfiles
            TextWriter fnd = new StreamWriter(onlineLog);
            TextWriter nfnd = new StreamWriter(notOnlineLog);
            fnd.WriteLine(DateTime.Now);
            fnd.WriteLine("online name,id,zcQTY,storeQTY,store name,id,ShouldUpdate");

            nfnd.WriteLine(DateTime.Now);
            nfnd.WriteLine("store name , store id");
            string connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + thefile + ";Extended Properties=\"Excel 8.0;HDR=YES;\"";
            DbProviderFactory factory = DbProviderFactories.GetFactory("System.Data.OleDb");
            // create connection to ZenCart
            string MyConString = "SERVER="+ConfigurationSettings.AppSettings["zcServer"]+";" +
                "DATABASE="+ConfigurationSettings.AppSettings["zcDatabase"]+";" +
                "UID=" + ConfigurationSettings.AppSettings["zcUser"] + ";" +
                "PASSWORD=" + ConfigurationSettings.AppSettings["zcPassword"] + ";";

            // these connection objects are for querying the zen cart database
            MySqlConnection mysqlCon = new MySqlConnection(MyConString);
            MySqlCommand mysqlCmd = mysqlCon.CreateCommand();
            MySqlDataReader Reader;
            mysqlCon.Open();

            // these connection objects are for updating the zen cart database
            MySqlConnection mysqlUpdateCon = new MySqlConnection(MyConString);
            MySqlCommand mysqlUpdateCmd = mysqlUpdateCon.CreateCommand();
            mysqlUpdateCon.Open();

            using (DbConnection connection = factory.CreateConnection())
            {
                connection.ConnectionString = connectionString;

                using (DbCommand command = connection.CreateCommand())
                {
                    command.CommandText = "SELECT * FROM [Sheet1$]";

                    connection.Open();

                    using (DbDataReader dr = command.ExecuteReader())
                    {


                        while (dr.Read())
                        {
                            // reading in one line from the spreadsheet
                            inStoreName = dr[1].ToString();
                            inStoreNumber = dr[0].ToString();
                            inStoreQTY = dr[2].ToString();
                            
                            Console.WriteLine(inStoreName);

                            // assign inStore names and ids and quantities to variables
                            // search for inStore ID in the online database
                            mysqlCmd.CommandText = "SELECT zen_products.products_id,zen_products.products_model, zen_products.products_quantity, zen_products_description.products_name "+
	                                                "FROM `zen_products` , `zen_products_description` "+
	                                                "WHERE zen_products.products_id = zen_products_description.products_id "+
	                                                "AND zen_products.products_model = " + inStoreNumber;
                            Reader = mysqlCmd.ExecuteReader();
                            if (!(Reader.HasRows))// if you don't find it online, then make a note of that and move to the next row in the dr.Read results
                            {
                                // LOG to nfnd
                                // LOG like this : store name , store Number
                                nfnd.WriteLine('"' + inStoreName + '"' + "," + inStoreNumber);
                            }
                            while (Reader.Read())
                            {
                                    
                                    zcName = Reader[(Reader.GetOrdinal("products_name"))].ToString();
                                    zcNumber = Reader[(Reader.GetOrdinal("products_model"))].ToString();
                                    zcQTY = Reader[(Reader.GetOrdinal("products_quantity"))].ToString();

                                    // LOG to fnd
                                    // LOG like this : online name , online model , online quantity , store name , store number , store quantity
                                    fnd.Write('"'+zcName +'"'+","+zcNumber+","+zcQTY+","+inStoreQTY+","+'"'+inStoreName+'"'+","+inStoreNumber);

                                    if (inStoreQTY != zcQTY)// if the quantities dont match, update the online data
                                    {
                                        // update query code goes here
                                        try
                                        {
                                            mysqlUpdateCmd.CommandText = "UPDATE zen_products SET products_quantity = '" + inStoreQTY +
                                                                    "' WHERE products_model = " + inStoreNumber;
                                           mysqlUpdateCmd.ExecuteNonQuery();
                                            // should probably trap that command with "try"
                                            fnd.WriteLine("," + mysqlUpdateCmd.CommandText);
                                        }
                                        catch (Exception ex) 
                                        {
                                           fnd.WriteLine (ex.Message.ToString()); 
                                        }
                                    }
                                    else
                                    {
                                        fnd.WriteLine(",No");
                                    }
                            }
                            Reader.Close();
                            
                        }
                    }
                }
                connection.Close();
            }
        // close file handles and database connections
            nfnd.Close();
            fnd.Close();
            mysqlCon.Close();

            // move the file to archive
            string archiveFileName = DTS + ".xls";
            string archivePath = System.IO.Path.Combine((System.Configuration.ConfigurationSettings.AppSettings["ArchivePath"]), archiveFileName);
            System.IO.File.Move(thefile,archivePath);
        }
    }
}
