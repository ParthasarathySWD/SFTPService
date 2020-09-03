using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Configuration;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.IO;
using System.Runtime.Caching;
using System.Net.Http;
using System.Net.Security;
using Newtonsoft.Json;
using System.Timers;
using System.Net;


namespace SFTPService
{
    public partial class Service1 : ServiceBase
    {
        private static System.Timers.Timer aTimer;

        public int TimerTick = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["TimerTick"]);

        public String FolderPath = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData) + @"\" + System.Configuration.ConfigurationManager.AppSettings["BaseFolderName"];
        public String Token = System.Configuration.ConfigurationManager.AppSettings["Token"];
        public String ActionURL = System.Configuration.ConfigurationManager.AppSettings["ActionURL"] + System.Configuration.ConfigurationManager.AppSettings["Token"];
        public String CompletedPath { get; set; }
        public String NewPath { get; set; }
        public String Error { get; set; }
        public String Execution { get; set; }
        public String Logs { get; set; }

        public string LogFile { get; set; }

        public Thread Worker = null;
        private MemoryCache _memCache;
        private CacheItemPolicy _cacheItemPolicy;
        private const int CacheTimeMilliseconds = 1000;

        public Service1()
        {
            InitializeComponent();

            CompletedPath = FolderPath + @"\Completed";
            Error = FolderPath + @"\Error";
            Execution = FolderPath + @"\Execution";
            NewPath = FolderPath + @"\New";
            Logs = FolderPath + @"\Logs";
            LogFile = Logs + "\\Logs_" +  DateTime.Now.ToString("MM-dd-yyyy") + ".txt";

        }

        protected override void OnStart(string[] args)
        {
            ThreadStart start = new ThreadStart(Working);
            Worker = new Thread(start);
            Worker.Start();
        }

        private void SetTimer()
        {
            // Create a timer with a two second interval.
            aTimer = new System.Timers.Timer(TimerTick);
            // Hook up the Elapsed event for the timer. 
            aTimer.Elapsed += OnTimedEvent;
            aTimer.AutoReset = false;
            aTimer.Enabled = true;
            aTimer.Stop();

        }

        private void OnTimedEvent(Object source, ElapsedEventArgs e)
        {
            StartExecution();
        }

        private void StopTimer()
        {
            aTimer.Stop();
        }

        private void StartTimer()
        {
            aTimer.Start();
        }

        public void Working()
        {
            _memCache = MemoryCache.Default;

            //Create all basic directory
            CreateBaseDirectory();

            SetTimer();

            FileSystemWatcher watcher = new FileSystemWatcher();
            watcher.Path = NewPath;
            WriteLog("Folder Path: " + FolderPath);

            // Watch both files and subdirectories.  
            watcher.IncludeSubdirectories = false;
            // Watch for all changes specified in the NotifyFilters  
            //enumeration.  
            watcher.NotifyFilter = NotifyFilters.LastWrite;
            // Watch all files.  
            watcher.Filter = "*.*";

            _cacheItemPolicy = new CacheItemPolicy()
            {
                RemovedCallback = OnRemovedFromCache
            };


            // Add event handlers.  
            watcher.Changed += new FileSystemEventHandler(OnChanged);

            //Start monitoring.  
            watcher.EnableRaisingEvents = true;

            WriteLog($"Service Started On " + DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss"));


        }

        private void StartExecution()
        {
            string timestamp = DateTime.Now.ToString("yyyyMMddHHmmssFFF");

            CreateBaseDirectoryToPath(CompletedPath + "\\" + timestamp);
            CreateBaseDirectoryToPath(Error + "\\" + timestamp);
            CreateBaseDirectoryToPath(Execution + "\\" + timestamp);

            MoveFilesFromNew_To_Execute(timestamp);
            string[] Files = Directory.GetFiles(Execution + "\\" + timestamp, "*.pdf");
            if (Files.Length != 0)
            {

                foreach (string file in Files)
                {
                    string filename = Path.GetFileName(file);
                    WriteLog("File Upload Started for " + filename);
                    bool response = UploadFile(file, filename);
                    if(response)
                    {
                        if (File.Exists(file))
                        {
                            File.Move(file, CompletedPath + "\\" + timestamp + "\\" + filename);
                        }
                    }
                    else
                    {
                        if (File.Exists(file))
                        {
                            File.Move(file, Error + "\\" + timestamp + "\\" + filename);
                        }
                    }
                }
            }
        }

        private void MoveFilesFromNew_To_Execute(String _timestamp)
        {
            if (Directory.Exists(NewPath))
            {
                string[] Files = Directory.GetFiles(NewPath, "*.pdf");

                if (Files.Length != 0)
                {

                    foreach (string file in Files)
                    {
                        string filename = Path.GetFileName(file);

                        //    if (!Directory.Exists(duplicate_order_directory + "\\" + DateTime.Now.ToString("yyyyMMdd") + "\\"))
                        //    {
                        //        Directory.CreateDirectory(duplicate_order_directory + "\\" + DateTime.Now.ToString("yyyyMMdd") + "\\");
                        //    }
                        //    if (File.Exists(duplicate_order_directory + "\\" + DateTime.Now.ToString("yyyyMMdd") + "\\" + croppedfilename + ".pdf"))
                        //    {
                        //        File.Delete(duplicate_order_directory + "\\" + DateTime.Now.ToString("yyyyMMdd") + "\\" + croppedfilename + ".pdf");
                        //    }
                        if(File.Exists(file))
                        {
                            File.Move(file, Execution + "\\" + _timestamp + "\\" + filename);
                        }

                        //}

                    }
                }
            }
        }

        //private async Task<System.IO.Stream> SendFile_To_WebService(FileStream _paramfs, string _FileName)
        //{

        //    HttpContent stringContent = new StringContent(Token);
        //    HttpContent fileStreamContent = new StreamContent(_paramfs);
        //    using (var client = new HttpClient())
        //    using (var formData = new MultipartFormDataContent())
        //    {
        //        formData.Add(stringContent, "Token", "param1");
        //        formData.Add(fileStreamContent, "image[]", _FileName);
        //        var response = await client.PostAsync(ActionURL, formData);
        //        if (!response.IsSuccessStatusCode)
        //        {
        //            return null;
        //        }
        //        return await response.Content.ReadAsStreamAsync();
        //    }
        //}

        private bool UploadFile(string filePath, string destFileName)
        {
            var uploaded = false;
            try
            {
                ServicePointManager.ServerCertificateValidationCallback += (sender, cert, chain, sslPolicyErrors) => true;

                ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;

                var httpClient = new HttpClient();
                var fileStream = File.Open(filePath, FileMode.Open);
                var fileInfo = new FileInfo(filePath);
                var content = new MultipartFormDataContent();
                content.Headers.Add("filePath", filePath);
                content.Add(new StreamContent(fileStream), "\"image[]\"", string.Format("\"{0}\"", destFileName));

                var task = httpClient.PostAsync(ActionURL, content).ContinueWith(t =>
                {
                    if (t.Status == TaskStatus.RanToCompletion)
                    {
                        var response = t.Result;
                        var jsonStringTask2 = response.Content.ReadAsStringAsync().ContinueWith(jsonStringResponse2 =>

                        {

                            string jsonString = Convert.ToString(jsonStringResponse2.Result.ToString());

                            // extract security token from results:
                            //dynamic securityTokenJSONObject = Newtonsoft.Json.Linq.JObject.Parse(jsonString);
                            //string securityToken = securityTokenJSONObject.access_token;
                            string logcontent = "";
                            if (jsonString == "Success")
                            {
                                uploaded = true;
                                logcontent = destFileName + " \\t Success \\t" + DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss");
                            }
                            else if(jsonString == "Failed")
                            {
                                uploaded = false;
                                logcontent = destFileName + " \\t Upload Failed \\t" + DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss");
                            }
                            else if (jsonString == "Order Not Found")
                            {
                                uploaded = false;
                                logcontent = destFileName + " \\t Order Not Found \\t" + DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss");
                            }
                            else if (jsonString == "File Size Exceeds Limit")
                            {
                                logcontent = destFileName + " \\t File Size Exceeds Limit \\t" + DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss");
                                uploaded = false;
                            }
                            else if(jsonString == "TokenID Not Valid")
                            {
                                logcontent = destFileName + " \\t TokenID Not Valid \\t" + DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss");
                                uploaded = false;
                            }
                            else
                            {
                                uploaded = false;
                                logcontent = destFileName + " \\t Server Error \\t" + DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss");
                            }

                            WriteLog(logcontent);
                        });

                        jsonStringTask2.Wait();
                    }
                    else
                    {
                        uploaded = false;
                        WriteLog(destFileName + " \\t Server Error \\t" + DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss"));
                    }

                    fileStream.Dispose();
                });

                task.Wait();
                httpClient.Dispose();
            }
            catch (Exception ex)
            {
                uploaded = false;
                throw ex;
            }

            return uploaded;
        }

        private void CreateBaseDirectory()
        {

            CreateBaseDirectoryToPath(CompletedPath);
            CreateBaseDirectoryToPath(NewPath);
            CreateBaseDirectoryToPath(Error);
            CreateBaseDirectoryToPath(Execution);
            CreateBaseDirectoryToPath(Logs);

        }

        private void CreateBaseDirectoryToPath(String Path)
        {

            if (!Directory.Exists(Path))
            {
                Directory.CreateDirectory(Path);

            }
        }

        // Handle cache item expiring 
        private void OnRemovedFromCache(CacheEntryRemovedArguments args)
        {
            if (args.RemovedReason != CacheEntryRemovedReason.Expired) return;

            var e = (FileSystemEventArgs)args.CacheItem.Value;

            StopTimer();
            StartTimer();
        }

        // Add file event to cache (won't add if already there so assured of only one occurance)
        private void OnChanged(object source, FileSystemEventArgs e)
        {
            _cacheItemPolicy.AbsoluteExpiration = DateTimeOffset.Now.AddMilliseconds(CacheTimeMilliseconds);
            _memCache.AddOrGetExisting(e.Name, e, _cacheItemPolicy);

        }

        public void WriteLog(String content)
        {
            // Specify what is done when a file is changed.  
            using (StreamWriter writer = new StreamWriter(LogFile, true))
            {
                writer.WriteLine(string.Format(content));
            }
        }

        protected override void OnStop()
        {
            try
            {
                if (Worker != null && Worker.IsAlive)
                {
                    Worker.Abort();
                }
            }
            catch (Exception)
            {

                throw;
            }
        }

        public int DeleteFile(string file)
        {
            File.Delete(file);
            return 1;
        }
    }
}
