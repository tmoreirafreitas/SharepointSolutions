using SPUtilitieServicesSolution.Extensions;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace SPUtilitieServicesSolution
{
    /// <summary>
    /// 
    /// </summary>
    public class SPLibraryService
    {
        /// <summary>
        /// 
        /// </summary>
        protected readonly ClientContext _ctx;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="ctx"></param>
        public SPLibraryService(ClientContext ctx)
        {
            _ctx = ctx;
        }

        /// <summary>
        /// Set LibraName and FileName and return the array of bytes.
        /// </summary>
        /// <param name="libraryName"></param>
        /// <param name="fileName"></param>
        /// <returns>The array of bytes</returns>
        /// <example>
        /// <code>
        /// var dataFile = DownloadFileRestFull("libraryName", "fileName");
        /// if(dataFile != null || dataFile.Length > 0)
        /// {
        ///     using (var file = new System.IO.FileStream("localPath", System.IO.FileMode.Create))
        ///     {
        ///         file.Write(dataFile, 0, dataFile.Length);
        ///     }
        /// }
        /// </code>
        /// </example>
        public byte[] DownloadFile(string libraryName, string fileName)
        {
            try
            {
                byte[] data = null;
                ICredentials credencial = _ctx.Credentials;
                var spList = _ctx.Web.Lists.GetByTitle(libraryName);
                _ctx.Load(spList, l => l.RootFolder, l => l.RootFolder.ServerRelativeUrl);
                _ctx.ExecuteQuery();

                if (spList != null)
                {
                    using (var client = new WebClient())
                    {
                        var uri = new Uri(_ctx.Url);
                        var path = string.Format("{0}/{1}/{2}", uri.GetLeftPart(UriPartial.Authority).TrimEnd('/'),
                            spList.RootFolder.ServerRelativeUrl.TrimStart('/'), fileName);
                        client.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
                        client.Headers.Add("User-Agent: Other");
                        client.Credentials = credencial;
                        data = client.DownloadData(path);
                    }
                }

                return data;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Set LibraName and FileName and return the List of FileNames.
        /// </summary>
        /// <param name="libraryName"></param>
        /// <param name="folderServerRelativeUrl"></param>
        /// <returns>Return the List of FileNames</returns>
        public IList<string> GetFileNames(string libraryName, string folderServerRelativeUrl)
        {
            try
            {
                var fileNames = new List<string>();
                var spList = _ctx.Web.Lists.GetByTitle(libraryName);
                _ctx.Load(spList, l => l.RootFolder, l => l.RootFolder.Folders);
                _ctx.ExecuteQuery();
                if (spList != null || spList.ItemCount > 0)
                {
                    CamlQuery caml = new CamlQuery();
                    var query = @"<View>
                                         <Query>
                                             <Where>
                                                  <Eq>
                                                      <FieldRef Name='FSObjType'/>
                                                      <Value Type='Lookup'>0</Value>
                                                  </Eq>
                                             </Where>
                                         </Query>
                                      </View>";

                    caml.ViewXml = query;
                    caml.FolderServerRelativeUrl = folderServerRelativeUrl;
                    var spItems = spList.GetItems(caml);
                    _ctx.Load(_ctx.Site);
                    _ctx.Load(spItems, items => items.Include(item => item.File,
                                                item => item.File.ServerRelativeUrl));
                    _ctx.ExecuteQuery();

                    if (spItems != null && spItems.Count > 0)
                    {
                        var uri = new Uri(_ctx.Site.Url);
                        object localLockObject = new object();

                        Parallel.ForEach(spItems,
                            new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount },
                            () => { return new List<string>(); },
                            (item, state, localList) =>
                            {
                                var newUrl = string.Format(@"{0}/{1}",
                                    uri.GetLeftPart(UriPartial.Authority).TrimEnd('/'), item.File.ServerRelativeUrl.TrimStart('/'));
                                localList.Add(newUrl);
                                return localList;
                            },
                        (finalResult) =>
                        {
                            lock (localLockObject) fileNames.AddRange(finalResult);
                        });
                    }
                }
                return fileNames;
            }
            catch (Exception ex)
            {
                //logger.Error(string.Format("StackTrace: {0}\nMessage: {1}", ex.StackTrace, ex.Message));
                throw ex;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="libraryName"></param>
        /// <returns></returns>
        public IDictionary<string, int> GetAllFolders(string libraryName)
        {
            try
            {
                var spFoldersMap = new Dictionary<string, int>();
                var spList = _ctx.Web.Lists.GetByTitle(libraryName);
                _ctx.Load(spList, l => l.RootFolder, l => l.RootFolder.Folders, l => l.ItemCount);
                _ctx.ExecuteQuery();

                if (spList != null && spList.ItemCount > 0)
                {
                    var camlToFolder = new CamlQuery();
                    var queryToFolder = @"<View Scope='RecursiveAll'>
                                                 <Query>
                                                     <Where>
                                                          <Eq>
                                                              <FieldRef Name='FSObjType'/>
                                                              <Value Type='Lookup'>1</Value>
                                                          </Eq>
                                                     </Where>
                                                 </Query>
                                              </View>";

                    camlToFolder.ViewXml = queryToFolder;

                    var camlToFile = new CamlQuery();
                    var queryToFiles = @"<View>
                                                 <Query>
                                                     <Where>
                                                          <Eq>
                                                              <FieldRef Name='FSObjType'/>
                                                              <Value Type='Lookup'>0</Value>
                                                          </Eq>
                                                     </Where>
                                                 </Query>
                                              </View>";

                    camlToFile.ViewXml = queryToFiles;
                    var spFiles = spList.GetItems(camlToFile);
                    var spFolders = spList.GetItems(camlToFolder);
                    _ctx.Load(spFolders);
                    _ctx.Load(spFiles);
                    _ctx.ExecuteQuery();

                    if (spFiles != null && spFiles.Count > 0)
                    {
                        var uri = new Uri(_ctx.Url);
                        var path = string.Format("{0}{1}", uri.GetLeftPart(UriPartial.Authority).TrimEnd('/'),
                            spList.RootFolder.ServerRelativeUrl);
                        spFoldersMap.Add(path, spFiles.Count);
                    }
                    if (spFolders != null && spFolders.Count > 0)
                    {
                        object localLockObject = new object();
                        Parallel.ForEach(spFolders,
                            new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount },
                            () => { return new Dictionary<string, int>(); },
                            (folder, state, localList) =>
                            {
                                try
                                {
                                    if (int.Parse(folder["ItemChildCount"].ToString()) > 0)
                                        localList.Add(folder["FileRef"].ToString(), int.Parse(folder["ItemChildCount"].ToString()));
                                }
                                catch (Exception ex)
                                {
                                    //logger.Error(string.Format("StackTrace: {0}\nMessage: {1}", ex.StackTrace, ex.Message));
                                    throw ex;
                                }
                                return localList;
                            },
                            (finalResult) =>
                            {
                                lock (localLockObject) spFoldersMap.ToList().AddRange(finalResult.ToList());
                            });
                    }
                }

                return spFoldersMap;
            }
            catch (Exception ex)
            {
                //logger.Error(string.Format("StackTrace: {0}\nMessage: {1}", ex.StackTrace, ex.Message));
                throw ex;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="listName"></param>
        /// <param name="fileUrl"></param>
        /// <param name="stream"></param>
        public void UploadFile(string listName, string fileUrl, System.IO.Stream stream)
        {
            var oList = _ctx.Web.Lists.GetByTitle(listName);
            if (oList != null)
            {
                using (var ms = new System.IO.MemoryStream())
                {
                    FileCreationInformation fci = new FileCreationInformation();
                    stream.CopyTo(ms);
                    fci.Content = ms.ToArray();
                    fci.Overwrite = true;
                    fci.Url = fileUrl;
                    _ctx.Load(_ctx.Web.RootFolder, f => f.Folders);
                    _ctx.ExecuteQuery();

                    var folders = _ctx.Web.RootFolder.Folders;
                    File file = oList.RootFolder.Files.Add(fci);
                    _ctx.ExecuteQuery();
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="fileUrl"></param>
        /// <param name="stream"></param>
        public void UploadFileRestFull(string fileUrl, System.IO.Stream stream)
        {
            var requestURI = new Uri(_ctx.Url);
            var baseDestURL = requestURI.GetLeftPart(UriPartial.Authority);
            var path = baseDestURL + fileUrl;
            byte[] buffer = new byte[stream.Length];

            try
            {
                using (var ms = new System.IO.MemoryStream())
                {
                    stream.CopyTo(ms);
                    buffer = ms.ToArray();
                }

                using (WebClient client = new WebClient())
                {
                    client.BaseAddress = baseDestURL;
                    client.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
                    client.Credentials = _ctx.Credentials;
                    var result = Encoding.UTF8.GetString(client.UploadData(path, "PUT", buffer));
                }
            }
            catch (OutOfMemoryException ex)
            {
                throw ex;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void CreateFolder(string listTitle, string folderName)
        {
            var list = _ctx.Web.Lists.GetByTitle(listTitle);
            var folderCreateInfo = new ListItemCreationInformation
            {
                UnderlyingObjectType = FileSystemObjectType.Folder,
                LeafName = folderName
            };
            var folderItem = list.AddItem(folderCreateInfo);
            folderItem.Update();
            _ctx.Web.Context.ExecuteQuery();
        }

        private void CreateSubFolderForFolder(string listName, string subfolderName, string folderName)
        {

            //This procedure creates a subfolder for a folder in a documentlibrary list
            var list = _ctx.Web.Lists.GetByTitle(listName);
            _ctx.Load(list, l => l.RootFolder, l => l.RootFolder.Folders);
            _ctx.ExecuteQuery();

            var path = list.RootFolder.ServerRelativeUrl.TrimEnd('/') + folderName;
            if (list != null)
            {
                ListItemCreationInformation newFolder = new ListItemCreationInformation();
                newFolder.UnderlyingObjectType = FileSystemObjectType.Folder;

                //This function gets the complete url to the folder where the subfolder is created for                
                _ctx.Load(_ctx.Web.RootFolder, rf => rf.Folders);
                _ctx.ExecuteQuery();
                newFolder.FolderUrl = path;
                newFolder.LeafName = subfolderName;
                ListItem item = list.AddItem(newFolder);
                item.Update();

                _ctx.ExecuteQuery();
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="libraryName"></param>
        /// <param name="relativeFolderPath"></param>
        public void CreateFolderSubFolderStructure(string libraryName, string relativeFolderPath)
        {
            string totalFolderName = String.Empty;
            string[] folders = relativeFolderPath.Split('/');
            var list = _ctx.Web.Lists.GetByTitle(libraryName);
            _ctx.Load(list, l => l.RootFolder, l => l.RootFolder.Folders);
            _ctx.ExecuteQuery();

            var path = list.RootFolder.ServerRelativeUrl.TrimEnd('/');

            foreach (string folderName in folders)
            {
                if (!string.IsNullOrEmpty(folderName))
                {
                    path += '/' + folderName;
                    Folder folder = null;
                    _ctx.Web.TryGetFileByServerRelativeUrl(path, out folder);
                    if (folder == null)
                    {
                        if (!String.IsNullOrEmpty(totalFolderName))
                            CreateSubFolderForFolder(libraryName, folderName, totalFolderName);
                        else
                            CreateFolder(libraryName, folderName);
                    }
                    totalFolderName += '/' + folderName;
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public bool ValidatePathSP(string path)
        {
            Folder hasFolder = null;
            _ctx.Web.TryGetFolderByServerRelativeUrl(path, out hasFolder);
            if (hasFolder != null)
                return true;
            else
                return false;
        }
    }
}
