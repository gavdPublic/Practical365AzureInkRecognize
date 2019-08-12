using Microsoft.SharePoint.Client;
using Microsoft.Win32;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Configuration;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Ink;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace Test_InkBasicConsoleApp
{
    // Add references to PresentationCore, PresentationFramework and WindowsBase
    // Add NuGets NewtonSoft.Json and Microsoft.SharePointOnline.CSOM
    // SharePoint Library with two fields: Single-Line-Of-Text and Picture
    // Configurations in the App.config file

    class Program : Application
    {
        const string azureEndpoint = "https://api.cognitive.microsoft.com";  // URL Azure Cognitive Service
        const string inkRecognitionUrl = "/inkrecognizer/v1.0-preview/recognize";  // URL ink recognition

        static Window myWindow;
        static InkCanvas myInkCanvas;

        [STAThread]
        static void Main(string[] args)
        {
            new Program().Run();  // Run the Canvas
            Tuple<string, string, string> FilePathAndName = GetInk();  // Call the Azure Ink Recognizer

            // Upload bmp file and text to SharePoint
            string BmpFileAndPath = FilePathAndName.Item1.Replace(".isf", "_Ink.bmp");
            string BmpFileName = FilePathAndName.Item2.Replace(".isf", "_Ink.bmp");
            string InkText = FilePathAndName.Item3;
            ClientContext spCtx = LoginCsom();
            UploadOneDocument(spCtx, BmpFileAndPath, BmpFileName);
            int docId = FindOneLibraryDoc(spCtx, BmpFileName);
            UpdateOneDoc(spCtx, docId, InkText, ConfigurationManager.AppSettings["spInkTextField"]);
            string BmpFileUrl = ConfigurationManager.AppSettings["spUrl"] + "/" +
                                ConfigurationManager.AppSettings["spLibrary"] + "/" + BmpFileName;
            UpdateOneDoc(spCtx, docId, BmpFileUrl, ConfigurationManager.AppSettings["spInkPictureField"]);
        }

        protected override void OnStartup(StartupEventArgs args)
        {
            base.OnStartup(args);

            myWindow = new Window();
            myInkCanvas = new InkCanvas();
            myWindow.Content = myInkCanvas;
            myWindow.Show();
        }

        static Tuple<string, string, string> GetInk()
        {
            SaveFileDialog mySaveFileDialog = new SaveFileDialog();
            mySaveFileDialog.Filter = "isf files (*.isf)|*.isf";

            if (mySaveFileDialog.ShowDialog() == true)
            {
                CreateIsfFile(mySaveFileDialog.FileName); // Save the ink to a .isf file
                CreateBmpFile(mySaveFileDialog.FileName); // Save the ink to a .bmp picture
                CreateJsonFile(mySaveFileDialog.FileName); // Save the strokes to a json file
            }

            string dataPath = mySaveFileDialog.FileName.Replace(".isf", "_Ink.json");
            var requestData = LoadJson(dataPath);
            string requestString = requestData.ToString(Formatting.None);
            string myText = RecognizeInk(requestString, mySaveFileDialog.FileName);

            Tuple<string, string, string> tplReturn = 
                new Tuple<string, string, string>(mySaveFileDialog.FileName, mySaveFileDialog.SafeFileName, myText);
            return tplReturn;
        }

        static void CreateIsfFile(string FilePathAndName)
        {
            string isfFileName = FilePathAndName.Replace(".isf", "_Ink.isf");
            using (FileStream isfFileStream = new FileStream(isfFileName, FileMode.Create))
            {
                myInkCanvas.Strokes.Save(isfFileStream);
            }
        }

        static void CreateBmpFile(string FilePathAndName)
        {
            int width = (int)myInkCanvas.ActualWidth;
            int height = (int)myInkCanvas.ActualHeight;
            RenderTargetBitmap myRenderBmp = new RenderTargetBitmap(width, height, 96d, 96d, PixelFormats.Default);
            myRenderBmp.Render(myInkCanvas);
            BmpBitmapEncoder myEncoder = new BmpBitmapEncoder();
            myEncoder.Frames.Add(BitmapFrame.Create(myRenderBmp));
            string bmpFileName = FilePathAndName.Replace(".isf", "_Ink.bmp");
            using (FileStream bmpFileStream = new FileStream(bmpFileName, FileMode.Create))
            {
                myEncoder.Save(bmpFileStream);
            }
        }

        static void CreateJsonFile(string FilePathAndName)
        {
            int intCounter = 1;
            string myStrokesJson = string.Empty;
            myStrokesJson = "{" +
                                "\"version\": 1, " +
                                "\"language\": \"en-US\", " +
                                "\"unit\": \"mm\", " +
                                "\"strokes\": [";
            foreach (Stroke oneStroke in myInkCanvas.Strokes)
            {
                string myPoints = string.Empty;
                foreach (Point onePoint in oneStroke.StylusPoints)
                {
                    myPoints += onePoint.X + "," + onePoint.Y + ",";
                }
                myPoints = myPoints.Remove(myPoints.Length - 1); // Remove last ","

                myStrokesJson += "{" +
                                    "\"id\": " + intCounter + "," +
                                    "\"points\": \"" +
                                    myPoints +
                                    "\"},";
                intCounter++;
            }
            myStrokesJson = myStrokesJson.Remove(myStrokesJson.Length - 1); // Remove last ","
            myStrokesJson += "]}";

            string jsonFileName = FilePathAndName.Replace(".isf", "_Ink.json");
            using (TextWriter writer = new StreamWriter(jsonFileName, true))
            {
                writer.Write(myStrokesJson);
            }
        }

        static void CreateResultJsonFile(string FilePathAndName, JObject JsonResult)
        {
            string myResult = JsonConvert.SerializeObject(JsonResult);

            string jsonFileName = FilePathAndName.Replace(".isf", "_Result.json");
            using (TextWriter writer = new StreamWriter(jsonFileName, true))
            {
                writer.Write(myResult);
            }
        }

        static void CreateResultTextFile(string FilePathAndName, string StringResult)
        {
            string jsonFileName = FilePathAndName.Replace(".isf", "_Result.txt");
            using (TextWriter writer = new StreamWriter(jsonFileName, true))
            {
                writer.Write(StringResult);
            }
        }

        static async Task<string> SendRequest(string apiAddress, string endpoint, string subscriptionKey, string requestData)
        {
            using (HttpClient client = new HttpClient { BaseAddress = new Uri(apiAddress) })
            {
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                client.DefaultRequestHeaders.Add("Ocp-Apim-Subscription-Key", subscriptionKey);

                var content = new StringContent(requestData, Encoding.UTF8, "application/json");
                var res = await client.PutAsync(endpoint, content);
                if (res.IsSuccessStatusCode)
                {
                    return await res.Content.ReadAsStringAsync();
                }
                else
                {
                    return $"ErrorCode: {res.StatusCode}";
                }
            }
        }

        static string RecognizeInk(string requestData, string FilePathAndName)
        {
            string subscriptionKey = ConfigurationManager.AppSettings["AzureSubscriptionKey"]; // Azure Subscription Key
            //construct the request
            var result = SendRequest(
                azureEndpoint,
                inkRecognitionUrl,
                subscriptionKey,
                requestData).Result;

            dynamic jsonObj = JsonConvert.DeserializeObject(result);
            CreateResultJsonFile(FilePathAndName, jsonObj);

            string myRecognizion = string.Empty;
            foreach (var oneUnit in jsonObj.recognitionUnits)
            {
                if (oneUnit.category == "line")
                {
                    myRecognizion += oneUnit.recognizedText + Environment.NewLine;
                }
            }

            CreateResultTextFile(FilePathAndName, myRecognizion);

            return myRecognizion;
        }

        public static JObject LoadJson(string FilePathAndName)
        {
            var jsonObj = new JObject();

            using (StreamReader file = System.IO.File.OpenText(FilePathAndName))
            using (JsonTextReader reader = new JsonTextReader(file))
            {
                jsonObj = (JObject)JToken.ReadFrom(reader);
            }
            return jsonObj;
        }

        static ClientContext LoginCsom()
        {
            ClientContext rtnContext = new ClientContext(
                ConfigurationManager.AppSettings["spUrl"]);

            SecureString securePw = new SecureString();
            foreach (
                char oneChar in ConfigurationManager.AppSettings["spUserPw"].ToCharArray())
            {
                securePw.AppendChar(oneChar);
            }
            rtnContext.Credentials = new SharePointOnlineCredentials(
                ConfigurationManager.AppSettings["spUserName"], securePw);

            return rtnContext;
        }

        static void UploadOneDocument(ClientContext spCtx, string FilePathAndName, string FileName)
        {
            List myList = spCtx.Web.Lists.GetByTitle(ConfigurationManager.AppSettings["spLibrary"]);

            using (FileStream myFileStream = new
                                    FileStream(FilePathAndName, FileMode.Open))
            {
                FileInfo myFileInfo = new FileInfo(FileName);
                spCtx.Load(myList.RootFolder);
                spCtx.ExecuteQuery();

                string fileUrl = String.Format("{0}/{1}",
                                myList.RootFolder.ServerRelativeUrl, myFileInfo.Name);
                Microsoft.SharePoint.Client.File.
                                SaveBinaryDirect(spCtx, fileUrl, myFileStream, true);
            }
        }

        static int FindOneLibraryDoc(ClientContext spCtx, string FileName)
        {
            List myList = spCtx.Web.Lists.GetByTitle(ConfigurationManager.AppSettings["spLibrary"]);

            int rowLimit = 10;
            string myViewXml = string.Format(@"
                <View>
                    <Query>
                        <Where>
                            <Eq>
                                <FieldRef Name='FileLeafRef' />
                                <Value Type='Text'>{0}</Value>
                            </Eq>
                        </Where>
                    </Query>
                    <ViewFields>
                        <FieldRef Name='FileLeafRef' />
                    </ViewFields>
                    <RowLimit>{1}</RowLimit>
                </View>", FileName, rowLimit);

            CamlQuery myCamlQuery = new CamlQuery();
            myCamlQuery.ViewXml = myViewXml;
            ListItemCollection allItems = myList.GetItems(myCamlQuery);
            spCtx.Load(allItems, itms => itms.Include(itm => itm["FileLeafRef"],
                                                     itm => itm.Id));
            spCtx.ExecuteQuery();

            return (int)allItems[0].Id;
        }

        static void UpdateOneDoc(ClientContext spCtx, int DocId, string DocText, string FieldName)
        {
            List myList = spCtx.Web.Lists.GetByTitle(ConfigurationManager.AppSettings["spLibrary"]);
            ListItem myListItem = myList.GetItemById(DocId);
            myListItem[FieldName] = DocText;

            myListItem.Update();
            spCtx.ExecuteQuery();
        }
    }
}
