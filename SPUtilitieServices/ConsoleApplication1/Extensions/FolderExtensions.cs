using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SPUtilitieServicesSolution.Extensions
{
    public static class FolderExtensions
    {
        /// <summary>
        /// Copy files 
        /// </summary>
        /// <param name="folder">Source Folder</param>
        /// <param name="folderUrl">Target Folder Url</param>
        public static void CopyFilesTo(this Folder folder, string folderUrl)
        {
            var ctx = (ClientContext)folder.Context;
            if (!ctx.Web.IsPropertyAvailable("ServerRelativeUrl"))
            {
                ctx.Load(ctx.Web, w => w.ServerRelativeUrl);
            }
            ctx.Load(folder, f => f.Files, f => f.ServerRelativeUrl, f => f.Folders);
            ctx.ExecuteQuery();

            //Ensure target folder exists
            ctx.Web.EnsureFolder(folderUrl.Replace(ctx.Web.ServerRelativeUrl, string.Empty));
            foreach (var file in folder.Files)
            {
                var targetFileUrl = file.ServerRelativeUrl.Replace(folder.ServerRelativeUrl, folderUrl);
                file.CopyTo(targetFileUrl, true);
            }
            ctx.ExecuteQuery();

            foreach (var subFolder in folder.Folders)
            {
                var targetFolderUrl = subFolder.ServerRelativeUrl.Replace(folder.ServerRelativeUrl, folderUrl);
                subFolder.CopyFilesTo(targetFolderUrl);
            }
        }
    }
}
