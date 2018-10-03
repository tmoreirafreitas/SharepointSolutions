using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SPUtilitieServicesSolution.Extensions
{
    public static class WebExtensions
    {
        /// <summary>
        /// Ensures whether the folder exists   
        /// </summary>
        /// <param name="web"></param>
        /// <param name="folderUrl"></param>
        /// <returns></returns>
        public static Folder EnsureFolder(this Web web, string folderUrl)
        {
            return EnsureFolderInternal(web.RootFolder, folderUrl);
        }


        private static Folder EnsureFolderInternal(Folder parentFolder, string folderUrl)
        {
            var ctx = parentFolder.Context;
            var folderNames = folderUrl.Split(new char[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
            var folderName = folderNames[0];
            var folder = parentFolder.Folders.Add(folderName);
            ctx.Load(folder);
            ctx.ExecuteQuery();

            if (folderNames.Length > 1)
            {
                var subFolderUrl = string.Join("/", folderNames, 1, folderNames.Length - 1);
                return EnsureFolderInternal(folder, subFolderUrl);
            }
            return folder;
        }

        public static bool TryGetFileByServerRelativeUrl(this Web web, string serverRelativeUrl, out Folder file)
        {
            var ctx = web.Context;
            try
            {
                var path = web.ServerRelativeUrl + serverRelativeUrl;
                file = web.GetFolderByServerRelativeUrl(serverRelativeUrl);
                ctx.Load(file);
                ctx.ExecuteQuery();
                return true;
            }
            catch (ServerException ex)
            {
                if (ex.ServerErrorTypeName == "System.IO.FileNotFoundException")
                {
                    file = null;
                    return false;
                }
                throw;
            }
        }

        public static bool TryGetFolderByServerRelativeUrl(this Web web, string serverRelativeUrl, out Folder folder)
        {
            var ctx = web.Context;
            try
            {
                var path = ctx.Url + serverRelativeUrl;
                folder = web.GetFolderByServerRelativeUrl(serverRelativeUrl);
                ctx.Load(folder);
                ctx.ExecuteQuery();
                return true;
            }
            catch (ServerException ex)
            {
                if (ex.ServerErrorTypeName == "System.IO.Microsoft.SharePoint.Client.FileNotFoundException")
                {
                    folder = null;
                    return false;
                }
                throw;
            }
        }
    }
}
