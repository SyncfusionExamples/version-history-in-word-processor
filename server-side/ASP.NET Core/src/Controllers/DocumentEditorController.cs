using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using System.Net.Http;
using System.Text;
using System.IO;
using Microsoft.AspNetCore.Cors;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Syncfusion.EJ2.DocumentEditor;
using WDocument = Syncfusion.DocIO.DLS.WordDocument;
using WFormatType = Syncfusion.DocIO.FormatType;
using Syncfusion.EJ2.SpellChecker;
using EJ2APIServices;
using SkiaSharp;
using BitMiracle.LibTiff.Classic;
using Newtonsoft.Json;
using EJ2APIServices_NET8.Hubs;
using Microsoft.AspNetCore.SignalR;
using Microsoft.Extensions.Configuration;
using Microsoft.EntityFrameworkCore;
using Microsoft.Data.SqlClient;
using System.Data;

namespace SyncfusionDocument.Controllers
{
    [Route("api/[controller]")]
    public class DocumentEditorController : Controller
    {
        private readonly IWebHostEnvironment _hostingEnvironment;
        string path;
        private readonly IHubContext<DocumentEditorHub> _hubContext;
        public static string connectionString;
        private static string fileLocation;
        private static byte saveThreshold = 200;


        public DocumentEditorController(IWebHostEnvironment hostingEnvironment , IHubContext<DocumentEditorHub> hubContext, IConfiguration config)
        {
            
            _hostingEnvironment = hostingEnvironment;
            path = Startup.path;
            _hubContext = hubContext;
            //Database connection string
            connectionString = config.GetConnectionString("DocumentEditorDatabase");
            fileLocation = _hostingEnvironment.WebRootPath;
        }
                     
        [AcceptVerbs("Post")]
        [HttpPost]
        [EnableCors("AllowAllOrigins")]
        [Route("Import")]
        public string Import(IFormCollection data)
        {
            if (data.Files.Count == 0)
                return null;
            Stream stream = new MemoryStream();
            IFormFile file = data.Files[0];
            int index = file.FileName.LastIndexOf('.');
            string type = index > -1 && index < file.FileName.Length - 1 ?
                file.FileName.Substring(index) : ".docx";
            file.CopyTo(stream);
            stream.Position = 0;

            //Hooks MetafileImageParsed event.
            WordDocument.MetafileImageParsed += OnMetafileImageParsed;
            WordDocument document = WordDocument.Load(stream, GetFormatType(type.ToLower()));
            //Unhooks MetafileImageParsed event.
            WordDocument.MetafileImageParsed -= OnMetafileImageParsed;

            string json = Newtonsoft.Json.JsonConvert.SerializeObject(document);
            document.Dispose();
            return json;
        }

        //Converts Metafile to raster image.
        private static void OnMetafileImageParsed(object sender, MetafileImageParsedEventArgs args)
        {
            if (args.IsMetafile)
            {
                //MetaFile image conversion(EMF and WMF)
                //You can write your own method definition for converting metafile to raster image using any third-party image converter.
                args.ImageStream = ConvertMetafileToRasterImage(args.MetafileStream);
            }
            else
            {
                //TIFF image conversion
                args.ImageStream = TiffToPNG(args.MetafileStream);

            }
        }

        // Converting Tiff to Png image using Bitmiracle https://www.nuget.org/packages/BitMiracle.LibTiff.NET
        private static MemoryStream TiffToPNG(Stream tiffStream)
        {
            MemoryStream imageStream = new MemoryStream();
            using (Tiff tif = Tiff.ClientOpen("in-memory", "r", tiffStream, new TiffStream()))
            {
                // Find the width and height of the image
                FieldValue[] value = tif.GetField(BitMiracle.LibTiff.Classic.TiffTag.IMAGEWIDTH);
                int width = value[0].ToInt();

                value = tif.GetField(BitMiracle.LibTiff.Classic.TiffTag.IMAGELENGTH);
                int height = value[0].ToInt();

                // Read the image into the memory buffer
                int[] raster = new int[height * width];
                if (!tif.ReadRGBAImage(width, height, raster))
                {
                    throw new Exception("Could not read image");
                }

                // Create a bitmap image using SkiaSharp.
                using (SKBitmap sKBitmap = new SKBitmap(width, height, SKImageInfo.PlatformColorType, SKAlphaType.Premul))
                {
                    // Convert a RGBA value to byte array.
                    byte[] bitmapData = new byte[sKBitmap.RowBytes * sKBitmap.Height];
                    for (int y = 0; y < sKBitmap.Height; y++)
                    {
                        int rasterOffset = y * sKBitmap.Width;
                        int bitsOffset = (sKBitmap.Height - y - 1) * sKBitmap.RowBytes;

                        for (int x = 0; x < sKBitmap.Width; x++)
                        {
                            int rgba = raster[rasterOffset++];
                            bitmapData[bitsOffset++] = (byte)((rgba >> 16) & 0xff);
                            bitmapData[bitsOffset++] = (byte)((rgba >> 8) & 0xff);
                            bitmapData[bitsOffset++] = (byte)(rgba & 0xff);
                            bitmapData[bitsOffset++] = (byte)((rgba >> 24) & 0xff);
                        }
                    }

                    // Convert a byte array to SKColor array.
                    SKColor[] sKColor = new SKColor[bitmapData.Length / 4];
                    int index = 0;
                    for (int i = 0; i < bitmapData.Length; i++)
                    {
                        sKColor[index] = new SKColor(bitmapData[i + 2], bitmapData[i + 1], bitmapData[i], bitmapData[i + 3]);
                        i += 3;
                        index += 1;
                    }

                    // Set the SKColor array to SKBitmap.
                    sKBitmap.Pixels = sKColor;

                    // Save the SKBitmap to PNG image stream.
                    sKBitmap.Encode(SKEncodedImageFormat.Png, 100).SaveTo(imageStream);
                    imageStream.Flush();
                }
            }
            return imageStream;
        }

        private static Stream ConvertMetafileToRasterImage(Stream ImageStream)
        {
            //Here we are loading a default raster image as fallback.
            Stream imgStream = GetManifestResourceStream("ImageNotFound.jpg");
            return imgStream;
            //To do : Write your own logic for converting metafile to raster image using any third-party image converter(Syncfusion doesn't provide any image converter).
        }

        private static Stream GetManifestResourceStream(string fileName)
        {
            System.Reflection.Assembly execAssembly = typeof(WDocument).Assembly;
            string[] resourceNames = execAssembly.GetManifestResourceNames();
            foreach (string resourceName in resourceNames)
            {
                if (resourceName.EndsWith("." + fileName))
                {
                    fileName = resourceName;
                    break;
                }
            }
            return execAssembly.GetManifestResourceStream(fileName);
        }

        [AcceptVerbs("Post")]
        [HttpPost]
        [EnableCors("AllowAllOrigins")]
        [Route("SpellCheck")]
        public string SpellCheck([FromBody] SpellCheckJsonData spellChecker)
        {
            try
            {
                SpellChecker spellCheck = new SpellChecker();
                spellCheck.GetSuggestions(spellChecker.LanguageID, spellChecker.TexttoCheck, spellChecker.CheckSpelling, spellChecker.CheckSuggestion, spellChecker.AddWord);
                return Newtonsoft.Json.JsonConvert.SerializeObject(spellCheck);
            }
            catch
            {
                return "{\"SpellCollection\":[],\"HasSpellingError\":false,\"Suggestions\":null}";
            }
        }

        [AcceptVerbs("Post")]
        [HttpPost]
        [EnableCors("AllowAllOrigins")]
        [Route("SpellCheckByPage")]
        public string SpellCheckByPage([FromBody] SpellCheckJsonData spellChecker)
        {
            try
            {
                SpellChecker spellCheck = new SpellChecker();
                spellCheck.CheckSpelling(spellChecker.LanguageID, spellChecker.TexttoCheck);
                return Newtonsoft.Json.JsonConvert.SerializeObject(spellCheck);
            }
            catch
            {
                return "{\"SpellCollection\":[],\"HasSpellingError\":false,\"Suggestions\":null}";
            }
        }

        public class SpellCheckJsonData
        {
            public int LanguageID { get; set; }
            public string TexttoCheck { get; set; }
            public bool CheckSpelling { get; set; }
            public bool CheckSuggestion { get; set; }
            public bool AddWord { get; set; }

        }

        public class UploadDocument
        {
            public string fileName { get; set; }
            public string documentOwner
            {
                get;
                set;
            }
        }
        
        public class CompareDocument
        {
            public string DocumentName { get; set; }
            public string SelectedVersion { get; set; }
        }

        [AcceptVerbs("Post")]
        [HttpPost]
        [EnableCors("AllowAllOrigins")]
        [Route("MailMerge")]
        public string MailMerge([FromBody] ExportData exportData)
        {
            Byte[] data = Convert.FromBase64String(exportData.documentData.Split(',')[1]);
            MemoryStream stream = new MemoryStream();
            stream.Write(data, 0, data.Length);
            stream.Position = 0;
            try
            {
                Syncfusion.DocIO.DLS.WordDocument document = new Syncfusion.DocIO.DLS.WordDocument(stream, Syncfusion.DocIO.FormatType.Docx);
                document.MailMerge.RemoveEmptyGroup = true;
                document.MailMerge.RemoveEmptyParagraphs = true;
                document.MailMerge.ClearFields = true;
                document.MailMerge.Execute(CustomerDataModel.GetAllRecords());
                document.Save(stream, Syncfusion.DocIO.FormatType.Docx);
            }
            catch (Exception ex)
            { }
            string sfdtText = "";
            Syncfusion.EJ2.DocumentEditor.WordDocument document1 = Syncfusion.EJ2.DocumentEditor.WordDocument.Load(stream, Syncfusion.EJ2.DocumentEditor.FormatType.Docx);
            sfdtText = Newtonsoft.Json.JsonConvert.SerializeObject(document1);
            document1.Dispose();
            return sfdtText;
        }
        public class CustomerDataModel
        {
            public static List<Customer> GetAllRecords()
            {
                List<Customer> customers = new List<Customer>();
                customers.Add(new Customer("9072379", "50%", "C/ Araquil, 67", "Madrid", "22020-08-10 00:00:00", "Spain", "Brittania", "2000", "19072379", "Folk och fä HB", "100000", "440", "32.34", "472.34", "28023", "12000", "2020-11-07 00:00:00", "2020-12-07 00:00:00"));
                customers.Add(new Customer("9072378", "20%", "C/ Araquil, 67", "Madrid", "22020-08-10 00:00:00", "Spain", "", "2", "19072369", "Maersk", "140000", "245", "20", "265", "28024", "12400", "2020-11-31 00:00:00", "2020-12-22300:00:00"));
                customers.Add(new Customer("9072377", "30%", "C/ Araquil, 67", "Madrid", "22020-08-10 00:00:00", "Spain", "Brittania", "100", "19072879", "Mediterranean Shipping Company", "104000", "434", "50.43", "484.43", "28025", "10000", "2020-11-07 00:00:00", "2020-12-02 00:00:00"));
                customers.Add(new Customer("9072393", "10%", "C/ Araquil, 67", "Madrid", "22020-08-10 00:00:00", "Spain", "Brittania", "2050", "19072378", "China Ocean Shipping Company", "175000", "500", "32", "532", "28026", "17000", "2020-09-23 00:00:00", "2020-10-09 00:00:00"));
                customers.Add(new Customer("9072377", "14%", "C/ Araquil, 67", "Madrid", "22020-08-10 00:00:00", "Spain", "Brittania", "2568", "19072380", "CGM", "155000", "655", "20.54", "675.54", "28027", "13000", "2020-10-11 00:00:00", "2020-11-17 00:00:00"));
                customers.Add(new Customer("9072376", "0%", "C/ Araquil, 67", "Madrid", "22020-08-10 00:00:00", "Spain", "Brittania", "1532", "19072345", " Hapag-Lloyd", "106500", "344", "30", "374", "28028", "14500", "2020-06-17 00:00:00", "2020-07-07 00:00:00"));
                customers.Add(new Customer("9072369", "05%", "C/ Araquil, 67", "Madrid", "22020-08-10 00:00:00", "Spain", "Brittania", "4462", "190723452", "Ocean Network Express", "100054", "541", "50", "591", "28029", "16500", "2020-04-07 00:00:00", "2020-05-07 00:00:00"));
                customers.Add(new Customer("9072359", "4%", "C/ Araquil, 67", "Madrid", "22020-08-10 00:00:00", "Spain", "Brittania", "27547", "190723713", "Evergreen Line", "124000", "800", "10.23", "810.23", "28030", "12500", "2020-03-07 00:00:00", "2020-04-07 00:00:00"));
                customers.Add(new Customer("9072380", "20%", "C/ Araquil, 67", "Madrid", "22020-08-10 00:00:00", "Spain", "Brittania", "7582", "19072312", "Yang Ming Marine Transport", "1046000", "290", "10", "300", "27631", "12670", "2020-11-10 00:00:00", "2020-12-13 00:00:00"));
                customers.Add(new Customer("9072381", "42%", "C/ Araquil, 67", "Madrid", "22020-08-10 00:00:00", "Spain", "Brittania", "862", "19072354", "Hyundai Merchant Marine", "145000", "800", "10.23", "810.23", "28032", "45000", "2020-10-17 00:00:00", "2020-12-23 00:00:00"));
                customers.Add(new Customer("9072391", "84%", "C/ Araquil, 67", "Madrid", "22020-08-10 00:00:00", "Spain", "Brittania", "82", "19072364", "Pacific International Line", "10094677", "344", "30", "374", "28033", "16500", "2020-11-14 00:00:00", "2020-12-21 00:00:00"));
                customers.Add(new Customer("9072392", "92%", "C/ Araquil, 67", "Madrid", "22020-08-10 00:00:00", "Spain", "Brittania", "82", "19072385", "Österreichischer Lloyd", "104270", "500", "32", "532", "28034", "156500", "2020-06-07 00:00:00", "2020-07-07 00:00:00"));
                return customers;
            }
        }
        public class Customer
        {
            public string CustomerID { get; set; }
            public string ProductName { get; set; }
            public string Quantity { get; set; }
            public string ShipName { get; set; }
            public string UnitPrice { get; set; }
            public string Discount { get; set; }
            public string ShipAddress { get; set; }
            public string ShipCity { get; set; }
            public string OrderDate { get; set; }
            public string ShipCountry { get; set; }
            public string OrderId { get; set; }
            public string Subtotal { get; set; }
            public string Freight { get; set; }
            public string Total { get; set; }
            public string ShipPostalCode { get; set; }
            public string RequiredDate { get; set; }
            public string ShippedDate { get; set; }
            public string ExtendedPrice { get; set; }
            public Customer(string orderId, string discount, string shipAddress, string shipCity, string orderDate, string shipCountry, string productName, string quantity, string customerID, string shipName, string unitPrice, string subtotal, string freight, string total, string shipPostalCode, string extendedPrice, string requiredDate, string shippedDate)
            {
                this.CustomerID = customerID;
                this.ProductName = productName;
                this.Quantity = quantity;
                this.ShipName = shipName;
                this.UnitPrice = unitPrice;
                this.Discount = discount;
                this.ShipAddress = shipAddress;
                this.ShipCity = shipCity;
                this.OrderDate = orderDate;
                this.ShipCountry = shipCountry;
                this.OrderId = orderId;
                this.Subtotal = subtotal;
                this.Freight = freight;
                this.Total = total;
                this.ShipPostalCode = shipPostalCode;
                this.ShippedDate = shippedDate;
                this.RequiredDate = requiredDate;
                this.ExtendedPrice = extendedPrice;
            }
        }
        public class ExportData
        {
            public string fileName { get; set; }
            public string modifiedUser { get; set; }
            public string documentData { get; set; }
        }




        public class CustomParameter
        {
            public string content { get; set; }
            public string type { get; set; }
        }

        [AcceptVerbs("Post")]
        [HttpPost]
        [EnableCors("AllowAllOrigins")]
        [Route("SystemClipboard")]
        public string SystemClipboard([FromBody] CustomParameter param)
        {
            if (param.content != null && param.content != "")
            {
                try
                {
                    //Hooks MetafileImageParsed event.
                    WordDocument.MetafileImageParsed += OnMetafileImageParsed;
                    WordDocument document = WordDocument.LoadString(param.content, GetFormatType(param.type.ToLower()));
                    //Unhooks MetafileImageParsed event.
                    WordDocument.MetafileImageParsed -= OnMetafileImageParsed;
                    string json = Newtonsoft.Json.JsonConvert.SerializeObject(document);
                    document.Dispose();
                    return json;
                }
                catch (Exception)
                {
                    return "";
                }
            }
            return "";
        }

        public class CustomRestrictParameter
        {
            public string passwordBase64 { get; set; }
            public string saltBase64 { get; set; }
            public int spinCount { get; set; }
        }
        [AcceptVerbs("Post")]
        [HttpPost]
        [EnableCors("AllowAllOrigins")]
        [Route("RestrictEditing")]
        public string[] RestrictEditing([FromBody] CustomRestrictParameter param)
        {
            if (param.passwordBase64 == "" && param.passwordBase64 == null)
                return null;
            return WordDocument.ComputeHash(param.passwordBase64, param.saltBase64, param.spinCount);
        }
        [AcceptVerbs("Post")]
        [HttpPost]
        [EnableCors("AllowAllOrigins")]
        [Route("LoadLatestVersionDocument")]
        public string   LoadLatestVersionDocument([FromBody] UploadDocument doc)
        {
            DocumentContent content = new DocumentContent();
            string[] fileEntries = System.IO.Directory.GetFiles("App_Data/" + doc.fileName + ".docx", "*.docx");
            DirectoryInfo directoryInfo = new DirectoryInfo("App_Data/" + doc.fileName + ".docx");

            // Get all files in the directory
            FileInfo[] files = directoryInfo.GetFiles();

            // Get the last modified file
            FileInfo lastModifiedFile = files
                .OrderByDescending(f => f.LastWriteTime)
                .FirstOrDefault();


            Stream stream = System.IO.File.OpenRead(Path.Combine("App_Data", doc.fileName + ".docx", lastModifiedFile.Name));
            stream.Position = 0;

            WordDocument document = WordDocument.Load(stream, FormatType.Docx);            
            if (doc.documentOwner != null)
            {
                int lastSyncedVersion = 0;
                List<ActionInfo> actions = CreatedTable(doc.fileName, out lastSyncedVersion);
                if (actions != null)
                {
                    //Updated pending edit from database to source document.
                    document.UpdateActions(actions);
                }
                string json = Newtonsoft.Json.JsonConvert.SerializeObject(document);
                content.version = lastSyncedVersion;
                content.sfdt = json;
                return Newtonsoft.Json.JsonConvert.SerializeObject(content);
            }
            else
            {
                document.Dispose();
                return Newtonsoft.Json.JsonConvert.SerializeObject(document);
            }
        }
        [HttpPost]
        [Route("UpdateAction")]
        [EnableCors("AllowAllOrigins")]
        public async Task<ActionInfo> UpdateAction([FromBody] ActionInfo param)
        {
            try
            {
                 ActionInfo modifiedAction = AddOperationsToTable(param);
                await _hubContext.Clients.Group(param.RoomName).SendAsync("dataReceived", "action", modifiedAction);
                return modifiedAction;
            }
            catch
            {
                return null;
            }
        }

        [HttpPost]
        [Route("GetActionsFromServer")]
        [EnableCors("AllowAllOrigins")]
        public string GetActionsFromServer([FromBody] ActionInfo param)
        {
            string tableName = param.RoomName;
            string getOperation = "SELECT * FROM \"" + tableName + "\" WHERE version > " + param.Version;
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                try
                {
                    SqlCommand command2 = new SqlCommand(getOperation, connection);
                    SqlCommand updateCommand = new SqlCommand(getOperation, connection);
                    connection.Open();
                    SqlDataReader reader = updateCommand.ExecuteReader();
                    DataTable table = new DataTable();
                    table.Load(reader);
                    DataTable oldTable = table;
                    if (table.Rows.Count > 0)
                    {
                        int startVersion = int.Parse(table.Rows[0]["version"].ToString());
                        int lowestVersion = GetLowestClientVersion(table);
                        if (startVersion > lowestVersion)
                        {
                            string updatedOperation = "SELECT * FROM \"" + tableName + "\" WHERE version >= " + lowestVersion;
                            SqlCommand command = new SqlCommand(updatedOperation, connection);
                            SqlDataReader reader2 = command.ExecuteReader();
                            table = new DataTable();
                            table.Load(reader2);
                        }
                        List<ActionInfo> actions = GetOperationsQueue(table);
                        foreach (ActionInfo info in actions)
                        {
                            if (!info.IsTransformed)
                            {
                                CollaborativeEditingHandler.TransformOperation(info, actions);
                            }
                        }
                        actions = actions.Where(x => x.Version > param.Version).ToList();
                        return Newtonsoft.Json.JsonConvert.SerializeObject(actions);
                    }
                }
                catch
                {
                    return "{}";
                }
            }
            return "{}";
        }

        public List<ActionInfo> CreatedTable(string roomName, out int lastSyncedVersion)
        {
            lastSyncedVersion = 0;
            string tableName = roomName;
            if (!TableExists(tableName))
            {

                string queryString = "CREATE TABLE \"" + tableName + "\" (version int IDENTITY(1,1) PRIMARY KEY, operation nvarchar(max), clientVersion int)";
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString, connection);
                    connection.Open();
                    command.ExecuteNonQuery();
                    // Create table to track the last saved version.
                    CreateRecordForVersionInfo(connection, roomName);

                }
            }
            else
            {

                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    connection.Open();
                    lastSyncedVersion = GetLastedSyncedVersion(connection, tableName);
                    string queryString = "SELECT * FROM \"" + tableName + "\" WHERE version > " + lastSyncedVersion;
                    SqlCommand command = new SqlCommand(queryString, connection);
                    SqlDataReader reader = command.ExecuteReader();
                    DataTable table = new DataTable();
                    table.Load(reader);
                    List<ActionInfo> actions = GetOperationsQueue(table);
                    return actions;

                }
            }
            return null;
        }
        public void CreateRecordForVersionInfo(SqlConnection connection, String roomName)
        {
            string tableName = "de_version_info";

            if (!TableExists(tableName))
            {
                // If table doesn't exist, create it
                string createTableQuery = $"CREATE TABLE \"{tableName}\" (roomName VARCHAR(MAX), lastSavedVersion INTEGER)";
                using (SqlCommand createTableCommand = new SqlCommand(createTableQuery, connection))
                {
                    createTableCommand.ExecuteNonQuery();
                }
            }

            // Insert record into the table
            string insertQuery = $"INSERT INTO \"{tableName}\" (roomName, lastSavedVersion) VALUES (@roomName, @lastSavedVersion)";
            using (SqlCommand insertCommand = new SqlCommand(insertQuery, connection))
            {
                insertCommand.Parameters.AddWithValue("@roomName", roomName);
                // Set initial version to 0
                insertCommand.Parameters.AddWithValue("@lastSavedVersion", 0);
                insertCommand.ExecuteNonQuery();
            }
            //}

        }
        private static bool TableExists(string tableName)
        {
            using (var connection = new SqlConnection(connectionString))
            {
                var command = new SqlCommand($"SELECT CASE WHEN OBJECT_ID('{tableName}', 'U') IS NOT NULL THEN 1 ELSE 0 END", connection);
                connection.Open();
                var result = (int)command.ExecuteScalar();               
                return result == 1;
            }
        }
        private ActionInfo AddOperationsToTable(ActionInfo action)
        {
            int clientVersion = action.Version;
            string tableName = action.RoomName;
            string value = Newtonsoft.Json.JsonConvert.SerializeObject(action);
            string query = "INSERT INTO \"" + tableName + "\" (operation, clientVersion) " + "VALUES (@Operation, @ClientVersion); ; SELECT SCOPE_IDENTITY() AS last_id";
            using (SqlConnection connection = new SqlConnection(connectionString))
            {

                SqlCommand command = new SqlCommand(query, connection);
                command.Parameters.Add("@Operation", SqlDbType.NVarChar).Value = value;
                command.Parameters.Add("@ClientVersion", SqlDbType.NVarChar).Value = action.Version;
                connection.Open();
                int updateVersion = int.Parse(command.ExecuteScalar().ToString());
                if (updateVersion - clientVersion == 1)
                {
                    action.Version = updateVersion;
                    UpdateCurrentActionToDB(tableName, action, connection);
                }
                else
                {
                    DataTable table = GetOperationsToTransform(tableName, clientVersion + 1, updateVersion, connection);
                    int startVersion = int.Parse(table.Rows[0]["version"].ToString());
                    int lowestVersion = GetLowestClientVersion(table);
                    if (startVersion > lowestVersion)
                    {
                        table = GetOperationsToTransform(tableName, lowestVersion, updateVersion, connection);
                    }
                    List<ActionInfo> actions = GetOperationsQueue(table);
                    foreach (ActionInfo info in actions)
                    {
                        if (!info.IsTransformed)
                        {
                            CollaborativeEditingHandler.TransformOperation(info, actions);
                        }
                    }
                    action = actions[actions.Count - 1];
                    action.Version = updateVersion;
                    UpdateCurrentActionToDB(tableName, actions[actions.Count - 1], connection);
                }
                if (updateVersion % saveThreshold == 0)
                {
                    UpdateOperationsToSourceDocument(tableName, HttpContext.Session.GetString("UserId"), true, updateVersion);
                }


            }
            return action;
        }
        private static int GetLowestClientVersion(DataTable table)
        {
            int clientVersion = int.Parse(table.Rows[0]["clientVersion"].ToString());
            foreach (DataRow row in table.Rows)
            {
                //TODO: Need to optimise version calculation for only untransformed operations
                int version = int.Parse(row["clientVersion"].ToString());
                if (version < clientVersion)
                {
                    clientVersion = version;
                }
            }
            return clientVersion;
        }
        private void UpdateCurrentActionToDB(string tableName, ActionInfo action, SqlConnection connection)
        {
            action.IsTransformed = true;
            string updateQuery = "UPDATE \"" + tableName + "\" SET operation = @Operation WHERE version = " + action.Version.ToString();
            SqlCommand updateCommand = new SqlCommand(updateQuery, connection);
            updateCommand.Parameters.Add("@Operation", SqlDbType.NVarChar).Value = Newtonsoft.Json.JsonConvert.SerializeObject(action);
            updateCommand.ExecuteNonQuery();
        }

        private static DataTable GetOperationsToTransform(string tableName, int clientVersion, int currentVersion, SqlConnection connection)
        {
            string getOperation = "SELECT * FROM \"" + tableName + "\" WHERE version BETWEEN " + clientVersion + " AND " + currentVersion.ToString();
            SqlCommand command = new SqlCommand(getOperation, connection);
            SqlDataReader reader = command.ExecuteReader();
            DataTable table = new DataTable();
            table.Load(reader);
            return table;
        }

        private static List<ActionInfo> GetOperationsQueue(DataTable table)
        {
            List<ActionInfo> actions = new List<ActionInfo>();
            foreach (DataRow row in table.Rows)
            {
                ActionInfo action = Newtonsoft.Json.JsonConvert.DeserializeObject<ActionInfo>(row["operation"].ToString());
                action.Version = int.Parse(row["version"].ToString());
                action.ClientVersion = int.Parse(row["clientVersion"].ToString());
                actions.Add(action);
            }
            return actions;
        }
        private static int GetLastedSyncedVersion(SqlConnection connection, string roomName)
        {
            string tableName = "de_version_info";
            string query = "SELECT lastSavedVersion FROM \"" + tableName + "\" WHERE roomName ='" + roomName + "'";
            var command = new SqlCommand(query, connection);
            command.Parameters.Add("@roomName", SqlDbType.NVarChar).Value = roomName;
            return int.Parse(command.ExecuteScalar().ToString());
        }

        static void UpdateModifiedVersion(string roomName, SqlConnection connection, int lastSavedVersion)
        {
            string tableName = "de_version_info";
            string query = "UPDATE [" + tableName + "] SET lastSavedVersion = @lastSavedVersion WHERE roomName = @roomName";
            using (SqlCommand command = new SqlCommand(query, connection))
            {

                command.Parameters.AddWithValue("@lastSavedVersion", lastSavedVersion);
                command.Parameters.AddWithValue("@roomName", roomName);
                command.ExecuteNonQuery();
            }
        }
        static void DeleteLastModifiedVersion(string roomName, SqlConnection connection)
        {
            string tableName = "de_version_info";
            string query = "DELETE FROM [" + tableName + "] WHERE roomName = @roomName";

            using (SqlCommand command = new SqlCommand(query, connection))
            {
                command.Parameters.AddWithValue("@roomName", roomName);
                command.ExecuteNonQuery();
            }
        }
        private static void DropTable(string documentId, SqlConnection connection)
        {
            try
            {
                //Delete operations record.
                string sqlQuery = "drop table \"" + documentId + "\"";
                var sqlCommand = new SqlCommand(sqlQuery, connection);
                sqlCommand.ExecuteNonQuery();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
        }

        public static void UpdateOperationsToSourceDocument(string fileName, string userId, bool partialSave, int endVersion)
        {
            try
            {
                SqlConnection connection = new SqlConnection(connectionString);
                connection.Open();
                string tableName = fileName;
                int lastSyncedVersion = GetLastedSyncedVersion(connection, fileName);
                string getOperation = "";
                if (partialSave)
                {
                    getOperation = "SELECT * FROM \"" + tableName + "\" WHERE version BETWEEN " + (lastSyncedVersion + 1).ToString() + " AND " + endVersion.ToString();
                    //getOperation = "SELECT Top (" + saveThreshold.ToString() + ") * FROM \"" + tableName + "\"";
                }
                else
                {
                    getOperation = "SELECT * FROM \"" + tableName + "\" WHERE version > " + lastSyncedVersion;
                }
                SqlCommand command = new SqlCommand(getOperation, connection);
                SqlDataReader reader = command.ExecuteReader();
                DataTable table = new DataTable();
                table.Load(reader);
                if (table.Rows.Count > 0)
                {
                    List<ActionInfo> actions = GetOperationsQueue(table);
                    foreach (ActionInfo info in actions)
                    {
                        if (!info.IsTransformed)
                        {
                            CollaborativeEditingHandler.TransformOperation(info, actions);
                        }
                    }
                    //CollaborativeEditingHandler handler = new CollaborativeEditingHandler(GetDocumentFromDatabase(fileName, GetSelectedDocumentOwner(userId, fileName, connection)));
                    var currentDirectory = System.IO.Directory.GetCurrentDirectory();
                    var outputdocName = fileName + ".docx";
                    int index = outputdocName.LastIndexOf('.');
                    string type = index > -1 && index < outputdocName.Length - 1 ?
                    outputdocName.Substring(index) : ".docx";

                    string[] fileEntries = System.IO.Directory.GetFiles("App_Data/" + outputdocName, "*.docx");
                    DirectoryInfo directoryInfo = new DirectoryInfo("App_Data/" + outputdocName);

                    // Get all files in the directory
                    FileInfo[] files = directoryInfo.GetFiles();

                    // Get the last modified file
                    FileInfo lastModifiedFile = files
                        .OrderByDescending(f => f.LastWriteTime)
                        .FirstOrDefault();

                    Stream stream1 = System.IO.File.OpenRead(Path.Combine("App_Data", outputdocName, lastModifiedFile.Name));

                    Syncfusion.EJ2.DocumentEditor.WordDocument document = Syncfusion.EJ2.DocumentEditor.WordDocument.Load(stream1, GetFormatType(type));
                    stream1.Close();
                    CollaborativeEditingHandler handler = new CollaborativeEditingHandler(document);
                    for (int i = 0; i < actions.Count; i++)
                    {
                        //Console.WriteLine(i);
                        handler.UpdateAction(actions[i]);
                    }
                    MemoryStream stream = new MemoryStream();
                    Syncfusion.DocIO.DLS.WordDocument doc = WordDocument.Save(Newtonsoft.Json.JsonConvert.SerializeObject(handler.Document));
                    doc.Save(stream, Syncfusion.DocIO.FormatType.Docx);
                    stream.Position = 0;
                    byte[] data = stream.ToArray();

                    string[] fileEntries1 = System.IO.Directory.GetFiles(Path.Combine("App_Data", outputdocName), "*.docx");
                    string filePath = Path.Combine("App_Data", outputdocName, string.Format("v{0}.docx", (fileEntries.Length + 1)));
                    using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                    {
                        fs.Write(data, 0, data.Length);
                    }

                    stream.Close();
                    if (!partialSave)
                    {
                        endVersion = actions[actions.Count - 1].Version;
                    }
                    doc.Close();
                }
                if (!partialSave)
                {
                    DeleteLastModifiedVersion(tableName, connection);
                    DropTable(fileName, connection);

                }
                else
                {
                    UpdateModifiedVersion(tableName, connection, endVersion);

                }

            }
            catch (Exception ex)
            {

            }

        }
        public class DocumentContent
        {
            public int version { get; set; }

            public string sfdt { get; set; }

            public List<ActionInfo> actions { get; set; } = new List<ActionInfo>();
        }
        public class SubChild
        {
            public string id { get; set; }
            public string name { get; set; }
            public string user { get; set; }
        }

        public class RootObject
        {
            public string id { get; set; }
            public string name { get; set; }
            public string user { get; set; }
            public bool expanded { get; set; }
            public List<SubChild> subChild { get; set; }
        }
        public class Version
        {
            public Version(string version, string fullName, string modifiedUser, DateTime? dateTime = null)
            {
                this.DocumentVersion = version;
                this.FullName = fullName;
                this.ModifiedUser = modifiedUser;
                this.LastSavedTime = dateTime ?? DateTime.Now;
            }
            public string DocumentVersion { get; set; }
            public string FullName { get; set; }
            public string ModifiedUser { get; set; }
            public DateTime LastSavedTime { get; set; }
        }
        public class CompareData
        {
            public string Document { get; set; }
            public List<RootObject> Data { get; set; }
        }
        //Compare the word document using DocIO https://help.syncfusion.com/file-formats/docio/word-document/compare-word-documents
        private WordDocument Compare(Version o, Version n)
        {
            using (FileStream originalDocumentStreamPath = new FileStream(o.FullName, FileMode.Open, FileAccess.Read))
            {
                using (WDocument originalDocument = new WDocument(originalDocumentStreamPath, WFormatType.Docx))
                {
                    //Load the revised document.
                    using (FileStream revisedDocumentStreamPath = new FileStream(n.FullName, FileMode.Open, FileAccess.Read))
                    {
                        using (WDocument revisedDocument = new WDocument(revisedDocumentStreamPath, WFormatType.Docx))
                        {
                            // Compare the original and revised Word documents.
                            originalDocument.Compare(revisedDocument);

                            return WordDocument.Load(originalDocument);
                        }
                    }
                }
            }
        }

        [AcceptVerbs("Post")]
        [HttpPost]
        [EnableCors("AllowAllOrigins")]
        [Route("CompareSelectedVersion")]
        public string CompareSelectedVersion([FromBody] CompareDocument doc)
        {
            string[] fileEntries = System.IO.Directory.GetFiles(("App_Data/" + doc.DocumentName), "*.docx");

            string targetFileName = doc.SelectedVersion;
            string previousFile = null;
            bool foundTargetFile = false;
            List<Version> versions = new List<Version>();
            string sfdtString = "";

            foreach (string filePath in fileEntries)
            {
                if (!foundTargetFile)
                {
                    if (System.IO.Path.GetFileName(filePath) == targetFileName)
                    {
                        versions.Add(new Version(System.IO.Path.GetFileName(filePath), filePath, null));
                        foundTargetFile = true;
                    }
                    else
                    {
                        previousFile = filePath; // Store the previous file
                    }
                }
                else
                {
                    break; // Exit the loop after finding the target file
                }
            }

            if (foundTargetFile && previousFile != null)
            {
                versions.Add(new Version(System.IO.Path.GetFileName(previousFile), previousFile, null, null));
                WordDocument result = Compare(versions[0], versions[1]);
                sfdtString = Newtonsoft.Json.JsonConvert.SerializeObject(result);
            }
            else
            {
                FileStream originalDocumentStreamPath = new FileStream(versions[0].FullName, FileMode.Open, FileAccess.Read);
                WordDocument document = WordDocument.Load(originalDocumentStreamPath, FormatType.Docx);
                sfdtString = Newtonsoft.Json.JsonConvert.SerializeObject(document);
                document.Dispose();
            }
            DocumentContent content = new DocumentContent();
            content.sfdt = sfdtString;
            return Newtonsoft.Json.JsonConvert.SerializeObject(content);
            // return sfdtString;
        }
        public class DownloadData
        {
            public string DocumentName { get; set; }
            public string SelectedVersion { get; set; }
        }
        [AcceptVerbs("Post")]
        [HttpPost]
        [EnableCors("AllowAllOrigins")]
        [Route("Download")]
        public IActionResult Download([FromBody] DownloadData data)
        {
            // Specify the file path relative to the App_Data folder
            string directoryPath = Path.Combine(_hostingEnvironment.ContentRootPath, "App_Data");
            string folderName = data.DocumentName;
            string versionFolder = data.SelectedVersion;
            string filePath = Path.Combine(directoryPath, folderName, versionFolder);

            // Check if the file exists
            if (!System.IO.File.Exists(filePath))
            {
                return NotFound(); // Return 404 Not Found if file doesn't exist
            }

            // Open the file stream
            FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            string fileName = data.DocumentName.Split('.')[0] + "_" + data.SelectedVersion;
            // Return FileStreamResult with the file stream
            return new FileStreamResult(fileStream, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            {
                FileDownloadName = fileName // Specify the file name to be used when downloaded
            };
        }

        [AcceptVerbs("Post")]
        [HttpPost]
        [EnableCors("AllowAllOrigins")]
        [Route("GetVersionData")]
        public string GetVersionData([FromBody] UploadDocument doc)
        {
            string directoryPath = "App_Data/" + doc.fileName;
            string[] fileEntries = System.IO.Directory.GetFiles("App_Data/" + doc.fileName, "*.docx");
            List<Version> versions = new List<Version>();
            // Check if the directory exists
            if (Directory.Exists(directoryPath))
            {
                // Get files in the directory
                string[] files = Directory.GetFiles(directoryPath);

                // Display file names with their last write time
                foreach (string file in files)
                {
                    string modifiedUser = "";
                    if (System.IO.File.Exists("App_Data\\fileNameWithUserName.json"))
                    {
                        // Read existing JSON data of file name with user
                        string existingJson = System.IO.File.ReadAllText("App_Data/fileNameWithUserName.json");

                        // Deserialize JSON data into a Dictionary<string, string>
                        Dictionary<string, string> data = JsonConvert.DeserializeObject<Dictionary<string, string>>(existingJson);

                        string fileToCheck = file.Replace("/", "\\");
                        // Check if the key of file exists in the dictionary
                        if (data.ContainsKey(fileToCheck))
                        {
                            // Key exists, retrieve its vale of user name
                            modifiedUser = data[fileToCheck];
                        }
                    }
                    // Get the last write time of the file
                    DateTime lastWriteTime = System.IO.File.GetLastWriteTime(file);
                    versions.Add(new Version(System.IO.Path.GetFileName(file), file, modifiedUser, lastWriteTime));
                }
                var versionDates = versions.OrderByDescending(v => v.LastSavedTime).ToList();
                CompareData compare = new CompareData();
                if (versions.Count > 1)
                {
                    // Compare versionDates[0] and versionDates[1]
                     compare.Document = Newtonsoft.Json.JsonConvert.SerializeObject(Compare(versionDates[0], versionDates[1]));
    
                }
                else
                {
                    // Handle the case where version[1] doesn't exist
                    FileStream originalDocumentStreamPath = new FileStream(versions[0].FullName, FileMode.Open, FileAccess.Read);
                    compare.Document = Newtonsoft.Json.JsonConvert.SerializeObject(WordDocument.Load(originalDocumentStreamPath, FormatType.Docx));
                    
                }
                var groupedDates = versions.GroupBy(version => version.LastSavedTime.Date).OrderByDescending(group => group.Key); ;
                // Iterate over grouped dates and create RootObject instances
                List<RootObject> rootObjects = new List<RootObject>();
                foreach (var group in groupedDates)
                {

                    List<SubChild> childObjects = new List<SubChild>();
                    List<Version> tempVersion = group.ToArray().OrderByDescending(v => v.LastSavedTime).ToList();
                    foreach (var chid in tempVersion)
                    {
                        SubChild childObject = new SubChild
                        {
                            id = chid.DocumentVersion,
                            name = chid.LastSavedTime.ToString("MMMM dd, hh:mm tt"),
                            user = chid.ModifiedUser

                        };
                        childObjects.Add(childObject);
                    }
                    // Create RootObject instance for each group
                    RootObject rootObject = new RootObject
                    {
                        id = group.Key.ToString("yyyy-MM-dd"),
                        name = group.Key.ToString("MMMM dd"),
                        user = childObjects[0].user,
                        subChild = childObjects
                    };

                    rootObjects.Add(rootObject);
                }
                compare.Data = rootObjects;
                return Newtonsoft.Json.JsonConvert.SerializeObject(compare);
            }
            return null;
        }
        private SubChild ParseSubChild()
        {
            return null;
        }
        [AcceptVerbs("Post")]
        [HttpPost]
        [EnableCors("AllowAllOrigins")]
        [Route("AutoSave")]
        public void AutoSave([FromBody] ExportData exportData)
        {
            byte[] data = Convert.FromBase64String(exportData.documentData.Split(',')[1]);
            string[] fileEntries = System.IO.Directory.GetFiles(Path.Combine("App_Data", exportData.fileName), "*.docx");
            string filePath = Path.Combine("App_Data", exportData.fileName, string.Format("v{0}.docx", (fileEntries.Length + 1)));
            using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                fs.Write(data, 0, data.Length);
            }
            // Read existing JSON data
            string jsonFilePath = Path.Combine("App_Data", "fileNameWithUserName.json");
            Dictionary<string, string> existingData;
            if (System.IO.File.Exists(jsonFilePath))
            {
                string existingJson = System.IO.File.ReadAllText(jsonFilePath);
                existingData = JsonConvert.DeserializeObject<Dictionary<string, string>>(existingJson);
            }
            else
            {
                existingData = new Dictionary<string, string>();
            }

            // Add the new data to the existing dictionary
            existingData[filePath] = exportData.modifiedUser;

            // Serialize the updated dictionary to JSON
            string updatedJson = JsonConvert.SerializeObject(existingData, Formatting.Indented);

            // Write the updated JSON data to the file
            System.IO.File.WriteAllText(jsonFilePath, updatedJson);
        }
        [AcceptVerbs("Post")]
        [HttpPost]
        [EnableCors("AllowAllOrigins")]
        [Route("LoadDefault")]
        public string LoadDefault()
        {
            Stream stream = System.IO.File.OpenRead("App_Data/GettingStarted.docx");
            stream.Position = 0;

            WordDocument document = WordDocument.Load(stream, FormatType.Docx);
            string json = Newtonsoft.Json.JsonConvert.SerializeObject(document);
            document.Dispose();
            return json;
        }

        [AcceptVerbs("Post")]
        [HttpPost]
        [EnableCors("AllowAllOrigins")]
        [Route("LoadDocument")]
        public string LoadDocument([FromForm] UploadDocument uploadDocument)
        {
            string documentPath = Path.Combine(path, uploadDocument.fileName);
            Stream stream = null;
            if (System.IO.File.Exists(documentPath))
            {
                byte[] bytes = System.IO.File.ReadAllBytes(documentPath);
                stream = new MemoryStream(bytes);
            }
            else
            {
                bool result = Uri.TryCreate(uploadDocument.fileName, UriKind.Absolute, out Uri uriResult)
                    && (uriResult.Scheme == Uri.UriSchemeHttp || uriResult.Scheme == Uri.UriSchemeHttps);
                if (result)
                {
                    stream = GetDocumentFromURL(uploadDocument.fileName).Result;
                    if (stream != null)
                        stream.Position = 0;
                }
            }
            WordDocument document = WordDocument.Load(stream, FormatType.Docx);
            string json = Newtonsoft.Json.JsonConvert.SerializeObject(document);
            document.Dispose();
            return json;
        }
        async Task<MemoryStream> GetDocumentFromURL(string url)
        {
            var client = new HttpClient(); ;
            var response = await client.GetAsync(url);
            var rawStream = await response.Content.ReadAsStreamAsync();
            if (response.IsSuccessStatusCode)
            {
                MemoryStream docStream = new MemoryStream();
                rawStream.CopyTo(docStream);
                return docStream;
            }
            else { return null; }
        }

        internal static FormatType GetFormatType(string format)
        {
            if (string.IsNullOrEmpty(format))
                throw new NotSupportedException("EJ2 DocumentEditor does not support this file format.");
            switch (format.ToLower())
            {
                case ".dotx":
                case ".docx":
                case ".docm":
                case ".dotm":
                    return FormatType.Docx;
                case ".dot":
                case ".doc":
                    return FormatType.Doc;
                case ".rtf":
                    return FormatType.Rtf;
                case ".txt":
                    return FormatType.Txt;
                case ".xml":
                    return FormatType.WordML;
                case ".html":
                    return FormatType.Html;
                default:
                    throw new NotSupportedException("EJ2 DocumentEditor does not support this file format.");
            }
        }
        internal static WFormatType GetWFormatType(string format)
        {
            if (string.IsNullOrEmpty(format))
                throw new NotSupportedException("EJ2 DocumentEditor does not support this file format.");
            switch (format.ToLower())
            {
                case ".dotx":
                    return WFormatType.Dotx;
                case ".docx":
                    return WFormatType.Docx;
                case ".docm":
                    return WFormatType.Docm;
                case ".dotm":
                    return WFormatType.Dotm;
                case ".dot":
                    return WFormatType.Dot;
                case ".doc":
                    return WFormatType.Doc;
                case ".rtf":
                    return WFormatType.Rtf;
                case ".html":
                    return WFormatType.Html;
                case ".txt":
                    return WFormatType.Txt;
                case ".xml":
                    return WFormatType.WordML;
                case ".odt":
                    return WFormatType.Odt;
                default:
                    throw new NotSupportedException("EJ2 DocumentEditor does not support this file format.");
            }
        }

        [AcceptVerbs("Post")]
        [HttpPost]
        [EnableCors("AllowAllOrigins")]
        [Route("Save")]
        public void Save([FromBody] SaveParameter data)
        {
            string name = data.FileName;
            string format = RetrieveFileType(name);
            if (string.IsNullOrEmpty(name))
            {
                name = "Document1.doc";
            }
            WDocument document = WordDocument.Save(data.Content);
            FileStream fileStream = new FileStream(name, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            document.Save(fileStream, GetWFormatType(format));
            document.Close();
            fileStream.Close();
        }

        [AcceptVerbs("Post")]
        [HttpPost]
        [EnableCors("AllowAllOrigins")]
        [Route("ExportSFDT")]
        public FileStreamResult ExportSFDT([FromBody] SaveParameter data)
        {
            string name = data.FileName;
            string format = RetrieveFileType(name);
            if (string.IsNullOrEmpty(name))
            {
                name = "Document1.doc";
            }
            WDocument document = WordDocument.Save(data.Content);
            return SaveDocument(document, format, name);
        }

        private string RetrieveFileType(string name)
        {
            int index = name.LastIndexOf('.');
            string format = index > -1 && index < name.Length - 1 ?
                name.Substring(index) : ".doc";
            return format;
        }

        public class SaveParameter
        {
            public string Content { get; set; }
            public string FileName { get; set; }
        }

        [AcceptVerbs("Post")]
        [HttpPost]
        [EnableCors("AllowAllOrigins")]
        [Route("Export")]
        public FileStreamResult Export(IFormCollection data)
        {
            if (data.Files.Count == 0)
                return null;
            string fileName = this.GetValue(data, "filename");
            string name = fileName;
            string format = RetrieveFileType(name);
            if (string.IsNullOrEmpty(name))
            {
                name = "Document1";
            }
            WDocument document = this.GetDocument(data);
            return SaveDocument(document, format, fileName);
        }

        private FileStreamResult SaveDocument(WDocument document, string format, string fileName)
        {
            Stream stream = new MemoryStream();
            string contentType = "";
            if (format == ".pdf")
            {
                contentType = "application/pdf";
            }
            else
            {
                WFormatType type = GetWFormatType(format);
                switch (type)
                {
                    case WFormatType.Rtf:
                        contentType = "application/rtf";
                        break;
                    case WFormatType.WordML:
                        contentType = "application/xml";
                        break;
                    case WFormatType.Html:
                        contentType = "application/html";
                        break;
                    case WFormatType.Dotx:
                        contentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.template";
                        break;
                    case WFormatType.Docx:
                        contentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
                        break;
                    case WFormatType.Doc:
                        contentType = "application/msword";
                        break;
                    case WFormatType.Dot:
                        contentType = "application/msword";
                        break;
                }
                document.Save(stream, type);
            }
            document.Close();
            stream.Position = 0;
            return new FileStreamResult(stream, contentType)
            {
                FileDownloadName = fileName
            };
        }

        private string GetValue(IFormCollection data, string key)
        {
            if (data.ContainsKey(key))
            {
                string[] values = data[key];
                if (values.Length > 0)
                {
                    return values[0];
                }
            }
            return "";
        }
        private WDocument GetDocument(IFormCollection data)
        {
            Stream stream = new MemoryStream();
            IFormFile file = data.Files[0];
            file.CopyTo(stream);
            stream.Position = 0;

            WDocument document = new WDocument(stream, WFormatType.Docx);
            stream.Dispose();
            return document;
        }
    }
}
