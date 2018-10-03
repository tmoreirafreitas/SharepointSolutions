using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Data;
using System.Dynamic;
using System.Linq;
using System.Linq.Expressions;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace SPUtilitieServicesSolution
{
    /// <summary>
    /// A classe SPListService,
    /// contém todos os métodos básicos para manipulação de CRUD nas listas
    /// </summary>
    public class SPListService
    {
        protected readonly ClientContext _ctx;

        public SPListService(ClientContext ctx)
        {
            _ctx = ctx;
        }

        public void AddNewListItem(string listName, IDictionary<string, string> keyValuePairs)
        {
            if (string.IsNullOrEmpty(listName))
            {
                throw new ArgumentException(string.Format("{0} is null or empty", listName), "listName");
            }

            try
            {
                var oList = _ctx.Web.Lists.GetByTitle(listName);
                _ctx.Load(oList);
                _ctx.Load(oList.Fields, fields => fields
                .Include(f => f.CanBeDeleted,
                f => f.InternalName, f => f.StaticName, f => f.FieldTypeKind, f => f.Title)
                .Where(f => f.FieldTypeKind != FieldType.Attachments && f.CanBeDeleted));
                _ctx.ExecuteQuery();
                if (oList != null)
                {
                    ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                    ListItem newItem = oList.AddItem(itemCreateInfo);

                    foreach (var key in keyValuePairs.Keys)
                    {
                        newItem[key] = keyValuePairs[key];
                    }

                    newItem.Update();
                    _ctx.ExecuteQuery();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void AddMultipleNewListItems(string listName, IList<IDictionary<string, string>> items)
        {
            if (string.IsNullOrEmpty(listName))
            {
                throw new ArgumentException(string.Format("{0} is null or empty", listName), "listName");
            }

            try
            {
                var oList = _ctx.Web.Lists.GetByTitle(listName);
                _ctx.Load(oList);
                _ctx.Load(oList.Fields, fields => fields
                .Include(f => f.CanBeDeleted,
                f => f.InternalName, f => f.StaticName, f => f.FieldTypeKind, f => f.Title)
                .Where(f => f.FieldTypeKind != FieldType.Attachments && f.CanBeDeleted));
                _ctx.ExecuteQuery();
                if (oList != null)
                {
                    var index = 0;
                    foreach (var item in items)
                    {
                        ++index;
                        ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                        ListItem newItem = oList.AddItem(itemCreateInfo);

                        foreach (var key in item.Keys)
                        {
                            newItem[key] = item[key];
                        }
                        newItem.Update();

                        if (index % 25 == 0)
                            _ctx.ExecuteQuery();
                    }
                    _ctx.ExecuteQuery();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public DataTable GetListsInfo()
        {
            try
            {
                var objs = new List<IDictionary<string, string>>();
                var oWeb = _ctx.Web;
                var oLists = oWeb.Lists;
                var query = (from oList in oLists
                             where (oList.BaseType == BaseType.GenericList || oList.BaseType == BaseType.DiscussionBoard ||
                             oList.BaseType == BaseType.Issue || oList.BaseType == BaseType.Survey)
                             select oList);

                var resultList = _ctx.LoadQuery(query);
                _ctx.ExecuteQuery();

                foreach (var list in resultList)
                {
                    var dic = new Dictionary<string, string>();
                    dic.Add("ID", list.Id.ToString());
                    dic.Add("NOME", list.Title);
                    dic.Add("QT_REGISTROS", list.ItemCount.ToString());
                    dic.Add("DT_CRIACAO", list.Created.ToString());
                    dic.Add("DT_MODIFICACAO", list.LastItemModifiedDate.ToString());
                    objs.Add(dic);
                }

                DataTable tb = Util.GetDataTableFromListDictionary(objs);
                return tb;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public DataTable GetListItemsByCaml(string listName, string query, string[] fieldNames = null)
        {
            if (string.IsNullOrEmpty(listName))
            {
                throw new ArgumentException("O nome da lista não pode ser vazia e nem nula", "listName");
            }
            if (string.IsNullOrEmpty(query))
            {
                throw new ArgumentException("A query não pode ser vazia e nem nula", "query");
            }

            try
            {
                var objs = new List<IDictionary<string, object>>();
                var oWeb = _ctx.Web;
                var oList = oWeb.Lists.GetByTitle(listName);
                _ctx.Load(oList);
                _ctx.Load(oList.Fields, fields => fields
                .Include(f => f.CanBeDeleted,
                f => f.InternalName, f => f.StaticName, f => f.FieldTypeKind, f => f.Title)
                .Where(f => f.FieldTypeKind != FieldType.Attachments && f.CanBeDeleted));
                _ctx.ExecuteQuery();

                ListItemCollectionPosition position = null;
                if (oList != null && oList.ItemCount > 0)
                {
                    var caml = new CamlQuery
                    {
                        ViewXml = query
                    };
                    do
                    {
                        caml.ListItemCollectionPosition = position;
                        var oListItems = oList.GetItems(caml);
                        if (fieldNames != null)
                        {
                            foreach (var fieldName in fieldNames)
                            {
                                _ctx.Load(oListItems, includes => includes.Include(i => i[fieldName]));
                            }
                            _ctx.Load(oListItems, inc => inc.ListItemCollectionPosition);
                        }
                        else
                            _ctx.Load(oListItems);
                        _ctx.ExecuteQuery();
                        position = oListItems.ListItemCollectionPosition;
                        foreach (var item in oListItems)
                        {
                            var spItem = new Dictionary<string, object>();

                            foreach (var key in item.FieldValues.Keys)
                            {
                                if (item.FieldValues[key] is FieldLookupValue LookupIdField)
                                {
                                    spItem.Add(key, string.Format("{0}#;{1}", LookupIdField.LookupId, LookupIdField.LookupValue));
                                }
                                else if (item.FieldValues[key] is FieldLookupValue[] FieldValues)
                                {
                                    var info = string.Empty;
                                    for (int i = 0; i < FieldValues.Length; i++)
                                    {
                                        info += string.Format("{0}#;{1}$", FieldValues[i].LookupId, FieldValues[i].LookupValue);
                                    }
                                    info.TrimEnd('$');
                                    spItem.Add(key, info);
                                }
                                else if (item.FieldValues[key] is FieldUserValue LookupUserIdField)
                                {
                                    spItem.Add(key, string.Format("{0}#;{1}", LookupUserIdField.LookupId, LookupUserIdField.LookupValue));
                                }
                                else if (item.FieldValues[key] is FieldUserValue[] FieldUsersValues)
                                {
                                    var info = string.Empty;
                                    for (int i = 0; i < FieldUsersValues.Length; i++)
                                    {
                                        info += string.Format("{0}#;{1}$", FieldUsersValues[i].LookupId, FieldUsersValues[i].LookupValue);
                                    }
                                    info.TrimEnd('$');
                                    spItem.Add(key, info);
                                }
                                else
                                {
                                    spItem.Add(key, item.FieldValues[key]);
                                }
                            }
                            objs.Add(spItem);
                        }
                    } while (position != null);
                }

                return Util.GetDataTableFromListDynamic(objs);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public DataTable GetAllListItem(string listName)
        {
            try
            {
                var objs = new List<IDictionary<string, object>>();
                var oWeb = _ctx.Web;
                var oList = oWeb.Lists.GetByTitle(listName);
                _ctx.Load(oList);
                _ctx.Load(oList.Fields, fields => fields
                .Include(f => f.CanBeDeleted,
                f => f.InternalName, f => f.StaticName, f => f.FieldTypeKind, f => f.Title)
                .Where(f => f.FieldTypeKind != FieldType.Attachments && f.CanBeDeleted));
                _ctx.ExecuteQuery();
                ListItemCollectionPosition position = null;
                if (oList != null && oList.ItemCount > 0)
                {
                    var caml = new CamlQuery();
                    var rowLimit = 1000;
                    caml.ViewXml = string.Format(@"<View Scope='RecursiveAll'><RowLimit Paged='TRUE'>{0}</RowLimit></View>", rowLimit);
                    do
                    {
                        caml.ListItemCollectionPosition = position;
                        var oListItems = oList.GetItems(caml);
                        _ctx.Load(oListItems);
                        _ctx.Load(oListItems, inc => inc.ListItemCollectionPosition);
                        _ctx.ExecuteQuery();
                        position = oListItems.ListItemCollectionPosition;

                        object listLock = new object();

                        Parallel.ForEach(oListItems, new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount }, 
                            () => { return new List<IDictionary<string, object>>(); },
                            (item, state, localList) =>
                            {
                                var spItem = new Dictionary<string, object>();

                                foreach (var key in item.FieldValues.Keys)
                                {
                                    if (item.FieldValues[key] is FieldLookupValue LookupIdField)
                                    {
                                        spItem.Add(key, string.Format("{0}#;{1}", LookupIdField.LookupId, LookupIdField.LookupValue));
                                    }
                                    else if (item.FieldValues[key] is FieldLookupValue[] FieldValues)
                                    {
                                        var info = string.Empty;
                                        for (int i = 0; i < FieldValues.Length; i++)
                                        {
                                            info += string.Format("{0}#;{1}$", FieldValues[i].LookupId, FieldValues[i].LookupValue);
                                        }
                                        info.TrimEnd('$');
                                        spItem.Add(key, info);
                                    }
                                    else if (item.FieldValues[key] is FieldUserValue LookupUserIdField)
                                    {
                                        spItem.Add(key, string.Format("{0}#;{1}", LookupUserIdField.LookupId, LookupUserIdField.LookupValue));
                                    }
                                    else if (item.FieldValues[key] is FieldUserValue[] FieldUsersValues)
                                    {
                                        var info = string.Empty;
                                        for (int i = 0; i < FieldUsersValues.Length; i++)
                                        {
                                            info += string.Format("{0}#;{1}$", FieldUsersValues[i].LookupId, FieldUsersValues[i].LookupValue);
                                        }
                                        info.TrimEnd('$');
                                        spItem.Add(key, info);
                                    }
                                    else
                                    {
                                        spItem.Add(key, item.FieldValues[key]);
                                    }
                                }

                                localList.Add(spItem);
                                return localList;
                            },
                            (finalResult) => { lock (listLock) objs.AddRange(finalResult); }
                        );

                    } while (position != null);
                }
                return Util.GetDataTableFromListDynamic(objs);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void UpdateListItemByCaml(string listName, string query, IDictionary<string, string> fieldsValues)
        {
            var oList = _ctx.Web.Lists.GetByTitle(listName);
            _ctx.Load(oList);
            _ctx.Load(oList.Fields, fields => fields.Include(f => f.CanBeDeleted)
                            .Where(f => f.FieldTypeKind != FieldType.Attachments && f.CanBeDeleted));
            _ctx.ExecuteQuery();

            var scopes = new Dictionary<int, ExceptionHandlingScope>();
            var infoErro = string.Empty;
            if (oList != null && oList.ItemCount > 0)
            {
                var caml = new CamlQuery();
                if (!string.IsNullOrEmpty(query))
                {
                    caml.ViewXml = query;
                }
                var spListItem = oList.GetItems(caml);
                _ctx.Load(spListItem);
                _ctx.ExecuteQuery();

                var count = 0;
                foreach (var item in spListItem)
                {
                    var scope = new ExceptionHandlingScope(_ctx);
                    scopes.Add(Convert.ToInt32(item["ID"].ToString()), scope);
                    using (scope.StartScope())
                    {
                        using (scope.StartTry())
                        {

                            ++count;
                            foreach (var f in fieldsValues.Keys)
                            {
                                if (item[f] != null)
                                {
                                    if (!string.IsNullOrEmpty(item[f].ToString()))
                                    {
                                        item[f] = fieldsValues[f];
                                    }
                                }
                            }
                            item.Update();
                        }

                        using (scope.StartCatch())
                        {

                        }
                    }
                    if (count % 25 == 0)
                        _ctx.ExecuteQuery();
                }

                _ctx.ExecuteQuery();
            }

            //try
            //{
            //    var oList = _ctx.Web.Lists.GetByTitle(listName);
            //    _ctx.Load(oList);
            //    _ctx.Load(oList.Fields, fields => fields.Include(f => f.CanBeDeleted)
            //                    .Where(f => f.FieldTypeKind != FieldType.Attachments && f.CanBeDeleted));
            //    _ctx.ExecuteQuery();
            //    if (oList != null && oList.ItemCount > 0)
            //    {
            //        var caml = new CamlQuery();
            //        if (!string.IsNullOrEmpty(query))
            //        {
            //            caml.ViewXml = query;
            //        }
            //        var spCListItem = oList.GetItems(caml);
            //        _ctx.Load(spCListItem);
            //        _ctx.ExecuteQuery();

            //        var count = 0;
            //        foreach (var item in spCListItem)
            //        {
            //            ++count;
            //            foreach (var f in fieldsValues.Keys)
            //            {
            //                if (item[f] != null)
            //                {
            //                    if (!string.IsNullOrEmpty(item[f].ToString()))
            //                    {
            //                        item[f] = fieldsValues[f];
            //                    }
            //                }
            //            }
            //            item.Update();
            //            if (count % 20 == 0)
            //                _ctx.ExecuteQuery();
            //        }

            //        _ctx.ExecuteQuery();
            //    }
            //}
            //catch (Exception ex)
            //{
            //    throw ex;
            //}
        }

        public void UpdateListItemById(string listName, string itemId, Dictionary<string, string> fieldsValues)
        {
            try
            {
                var oList = _ctx.Web.Lists.GetByTitle(listName);
                _ctx.Load(oList);
                _ctx.Load(oList.Fields, fields => fields.Include(f => f.CanBeDeleted)
                                .Where(f => f.FieldTypeKind != FieldType.Attachments && f.CanBeDeleted));
                _ctx.ExecuteQuery();
                if (oList != null && oList.ItemCount > 0)
                {
                    var item = oList.GetItemById(itemId);
                    _ctx.Load(item);
                    _ctx.ExecuteQuery();

                    foreach (var f in fieldsValues.Keys)
                        item[f] = fieldsValues[f];

                    item.Update();
                    _ctx.ExecuteQuery();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public bool DeleteListItemById(string listName, int itemId)
        {
            var isDeleted = false;
            try
            {
                if (string.IsNullOrEmpty(listName))
                {
                    throw new ArgumentException(string.Format("{0} is null or empty", listName), "listName");
                }
                var oList = _ctx.Web.Lists.GetByTitle(listName);
                _ctx.Load(oList);
                _ctx.Load(oList.Fields, fields => fields
                .Include(f => f.CanBeDeleted,
                f => f.InternalName, f => f.StaticName, f => f.FieldTypeKind, f => f.Title)
                .Where(f => f.FieldTypeKind != FieldType.Attachments && f.CanBeDeleted));
                _ctx.ExecuteQuery();
                if (oList != null && oList.ItemCount > 0)
                {
                    var oListItem = oList.GetItemById(itemId);
                    oListItem.DeleteObject();
                    _ctx.ExecuteQuery();
                    isDeleted = true;
                }
                else
                {
                    isDeleted = false;
                }
                return isDeleted;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void DeleteListItems(string listName, int[] ids, IProgress<string> progress = null) // Action<string, long, long> reportProgress = null,
        {
            if (string.IsNullOrEmpty(listName))
            {
                throw new ArgumentException(string.Format("{0} is null or empty", listName), "listName");
            }
            if (ids == null)
            {
                throw new ArgumentException("the ids is null", "ids");
            }

            var scopes = new Dictionary<int, ExceptionHandlingScope>();
            var oList = _ctx.Web.Lists.GetByTitle(listName);
            _ctx.Load(oList);
            _ctx.Load(oList.Fields, fields => fields
            .Include(f => f.CanBeDeleted,
            f => f.InternalName, f => f.StaticName, f => f.FieldTypeKind, f => f.Title)
            .Where(f => f.FieldTypeKind != FieldType.Attachments && f.CanBeDeleted));
            _ctx.ExecuteQuery();
            if (oList != null && oList.ItemCount > 0)
            {
                var index = 0;

                foreach (var id in ids)
                {
                    ExceptionHandlingScope scope = new ExceptionHandlingScope(_ctx);
                    scopes.Add(id, scope);
                    using (scope.StartScope())
                    {
                        using (scope.StartTry())
                        {
                            index++;
                            var oListItem = oList.GetItemById(id);
                            oListItem.DeleteObject();
                        }
                        using (scope.StartCatch())
                        {

                        }
                    }
                    if (index % 100 == 0)
                    {
                        _ctx.ExecuteQuery();

                        if (progress != null)
                        {
                            string message = string.Format("Lista:{0}\nPor favor! Aguarde enquanto os registros são deletados.\n{1} registros deletados de {2} ", oList.Title, index, oList.ItemCount);
                            progress.Report(message);
                        }
                    }
                }
                _ctx.ExecuteQuery();

                var query = scopes.Where(s => s.Value.ServerErrorCode != -1);
                foreach (var q in query)
                {
                    if (progress != null)
                    {
                        string message = string.Format("Um erro ocorreu no item ({0})\r\nErrorCode:{1}\r\nError:{2}", q.Key, q.Value.ServerErrorCode, q.Value.ErrorMessage);
                        progress.Report(message);
                    }
                }
            }
            //try
            //{
            //    var oList = _ctx.Web.Lists.GetByTitle(listName); 
            //    _ctx.Load(oList);
            //    _ctx.Load(oList.Fields, fields => fields
            //    .Include(f => f.CanBeDeleted,
            //    f => f.InternalName, f => f.StaticName, f => f.FieldTypeKind, f => f.Title)
            //    .Where(f => f.FieldTypeKind != FieldType.Attachments && f.CanBeDeleted));
            //    _ctx.ExecuteQuery();
            //    if (oList != null && oList.ItemCount > 0)
            //    {
            //        var index = 0;
            //        foreach (var id in ids)
            //        {
            //            index++;
            //            var oListItem = oList.GetItemById(id);
            //            oListItem.DeleteObject();
            //            if (index % 25 == 0)
            //            {
            //                _ctx.ExecuteQuery();

            //                if (progress != null)
            //                {
            //                    string message = string.Format("Lista:{0}\nPor favor! Aguarde enquanto os registros são deletados.\n{1} registros deletados de {2} ", oList.Title, index, oList.ItemCount);
            //                    progress.Report(message);
            //                }
            //            }
            //        }
            //        _ctx.ExecuteQuery();
            //    }
            //}
            //catch (Exception ex)
            //{
            //    throw ex;
            //}
        }

        public async Task DeleteListItemsAsync(string listName, string query = null, IProgress<string> progress = null)
        {
            await Task.Run(() =>
             {
                 ListItemCollectionPosition licp = null;
                 var oList = _ctx.Web.Lists.GetByTitle(listName);
                 _ctx.Load(oList);
                 _ctx.Load(oList.Fields, fields => fields
                 .Include(f => f.CanBeDeleted,
                 f => f.InternalName, f => f.StaticName, f => f.FieldTypeKind, f => f.Title)
                 .Where(f => f.FieldTypeKind != FieldType.Attachments && f.CanBeDeleted));
                 _ctx.ExecuteQuery();

                 var scopes = new Dictionary<int, ExceptionHandlingScope>();
                 if (oList != null && oList.ItemCount > 0)
                 {
                     var itemCount = oList.ItemCount;
                     var currentCount = 0;
                     do
                     {
                         var camlQuery = new CamlQuery();
                         camlQuery.ViewXml = string.IsNullOrEmpty(query) ? @"<View><ViewFields><FieldRef Name='ID'/></ViewFields><RowLimit>100</RowLimit></View>" : query;
                         camlQuery.ListItemCollectionPosition = licp;
                         var items = oList.GetItems(camlQuery);
                         _ctx.Load(items, its => its.Include(item => item["ID"], item => item["Created"], item => item["Modified"]));
                         _ctx.ExecuteQuery();
                         licp = items.ListItemCollectionPosition;
                         currentCount += items.Count;

                         foreach (var item in items.ToList())
                         {
                             var scope = new ExceptionHandlingScope(_ctx);
                             scopes.Add(Convert.ToInt32(item["ID"].ToString()), scope);
                             using (scope.StartScope())
                             {
                                 using (scope.StartTry())
                                 {
                                     item.DeleteObject();
                                 }
                                 using (scope.StartCatch())
                                 {

                                 }
                             }
                             if (progress != null)
                             {
                                 string message = string.Format("Lista:{0}\nPor favor! Aguarde enquanto os registros são deletados.\n{1} registros deletados de {2} ", oList.Title, currentCount, itemCount);
                                 progress.Report(message);
                             }
                         }
                         _ctx.ExecuteQuery();

                         var erros = scopes.Where(s => s.Value.ServerErrorCode != -1);
                         foreach (var erro in erros)
                         {
                             if (progress != null)
                             {
                                 string message = string.Format("Um erro ocorreu no item ({0})\r\nErrorCode:{1}\r\nError:{2}", erro.Key, erro.Value.ServerErrorCode, erro.Value.ErrorMessage);
                                 progress.Report(message);
                             }
                         }

                     } while (licp != null);
                 }
                 //try
                 //{
                 //    ListItemCollectionPosition licp = null;
                 //    var oList = _ctx.Web.Lists.GetByTitle(listName);
                 //    _ctx.Load(oList);
                 //    _ctx.Load(oList.Fields, fields => fields
                 //    .Include(f => f.CanBeDeleted,
                 //    f => f.InternalName, f => f.StaticName, f => f.FieldTypeKind, f => f.Title)
                 //    .Where(f => f.FieldTypeKind != FieldType.Attachments && f.CanBeDeleted));
                 //    _ctx.ExecuteQuery();

                 //    if (oList != null && oList.ItemCount > 0)
                 //    {
                 //        var itemCount = oList.ItemCount;
                 //        var currentCount = 0;
                 //        do
                 //        {
                 //            var camlQuery = new CamlQuery();
                 //            camlQuery.ViewXml = string.IsNullOrEmpty(query) ? @"<View><ViewFields><FieldRef Name='ID'/></ViewFields><RowLimit>100</RowLimit></View>" : query;
                 //            camlQuery.ListItemCollectionPosition = licp;
                 //            var items = oList.GetItems(camlQuery);
                 //            _ctx.Load(items, its => its.Include(item => item["ID"], item => item["Created"], item => item["Modified"]));
                 //            _ctx.ExecuteQuery();
                 //            licp = items.ListItemCollectionPosition;
                 //            currentCount += items.Count;
                 //            foreach (var item in items.ToList())
                 //            {
                 //                item.DeleteObject();
                 //            }
                 //            _ctx.ExecuteQuery();

                 //            if (progress != null)
                 //            {
                 //                string message = string.Format("Lista:{0}\nPor favor! Aguarde enquanto os registros são deletados.\n{1} registros deletados de {2} ", oList.Title, currentCount, itemCount);
                 //                progress.Report(message);
                 //            }
                 //        } while (licp != null);
                 //    }
                 //}
                 //catch (Exception ex)
                 //{
                 //    throw ex;
                 //}
             });
        }

        public IEnumerable<IDictionary<string, byte[]>> DownloadAttachedFileFromListItem(string listName, int itemId)
        {
            var listFiles = new List<IDictionary<string, byte[]>>();
            var dicFiles = new Dictionary<string, byte[]>();
            var scope = new ExceptionHandlingScope(_ctx);

            using (scope.StartScope())
            {
                using (scope.StartTry())
                {
                    var oWeb = _ctx.Web;
                    var oList = oWeb.Lists.GetByTitle(listName);
                    var oItem = oList.GetItemById(itemId);
                    _ctx.Load(oList, l => l.RootFolder.Folders.Where(f => f.Name == "Attachments"));
                    if (oList.RootFolder.Folders.Count > 0)
                    {
                        Folder attachmentFolder = oList.RootFolder.Folders[0];
                        _ctx.Load(attachmentFolder, f => f.Folders);

                        foreach (Folder itemFolder in attachmentFolder.Folders)
                        {
                            FileCollection files = itemFolder.Files;
                            _ctx.Load(files, fls => fls.Include(f => f.ServerRelativeUrl, f => f.Name));

                            if (files.Count > 0)
                            {
                                for (int i = 0; i < files.Count; i++)
                                {
                                    using (var ms = new System.IO.MemoryStream())
                                    {
                                        using (var fileInformation = File.OpenBinaryDirect(_ctx, files[i].ServerRelativeUrl))
                                        {
                                            fileInformation.Stream.CopyTo(ms);
                                            dicFiles.Add(files[i].Name, ms.ToArray());
                                            listFiles.Add(dicFiles);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                using (scope.StartCatch())
                {

                }
            }

            if (scope.HasException)
            {
                string message = string.Format("Um erro ocorreu na lista ({0})\r\nErrorCode:{1}\r\nError:{2}\r\nStackTrace{3}",
                    listName, scope.ServerErrorCode, scope.ErrorMessage, scope.ServerStackTrace);
                throw new Exception(message);
            }
            _ctx.ExecuteQuery();
            return listFiles;
        }

        public void AttachFileToListItem(string listName, int itemId, string fileName, bool overwrite)
        {
            using (var fileStream = new System.IO.FileStream(fileName, System.IO.FileMode.Open))
            {
                string attachmentPath = string.Format("/{0}/Lists/{1}/Attachments/{2}/{3}",
                    _ctx.Web.ServerRelativeUrl, listName, itemId, System.IO.Path.GetFileName(fileName));
                File.SaveBinaryDirect(_ctx, attachmentPath, fileStream, overwrite);
            }
        }

        public void DeleteAttachedFileFromListItem(string listName, int itemId, string attachmentFileName)
        {
            //http://siteurl/lists/[listname]/attachments/[itemid]/[filename]
            string attachmentPath = string.Format("/{0}/lists/{1}/Attachments/{2}/{3}", _ctx.Web.ServerRelativeUrl,
                listName, itemId, System.IO.Path.GetFileName(attachmentFileName));
            var file = _ctx.Web.GetFileByServerRelativeUrl(attachmentPath);
            file.DeleteObject();
            _ctx.ExecuteQuery();
        }
    }
}
