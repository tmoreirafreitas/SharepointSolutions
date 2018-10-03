using Microsoft.SharePoint.Client;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Client.Utilities;
using Microsoft.SharePoint.Client.Application;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.Dynamic;
using System.Web;

namespace SPUtilitieServicesSolution
{
    public class SPUtilitieServices : IDisposable
    {
        private string _url = string.Empty;
        private bool disposed = false;
        private ICredentials _credencial = null;
        private ClientContext _ctx = null;
        private SPGroupService spGroups;
        private SPUserService spUsers;
        private SPListService spList;
        private SPLibraryService spls;

        public SPUtilitieServices(string siteURL, ICredentials credencial)
        {

            if (credencial == null)
            {
                throw new ArgumentException(string.Format("{0} is null or empty", credencial), "SiteCollection");
            }
            else
            {
                if (!string.IsNullOrEmpty(siteURL))
                {
                    _url = siteURL;
                    _credencial = credencial;
                    _ctx = new ClientContext(_url)
                    {
                        Credentials = _credencial,
                        RequestTimeout = System.Threading.Timeout.Infinite
                    };
                    _ctx.PendingRequest.RequestExecutor.RequestKeepAlive = false;
                    _ctx.PendingRequest.RequestExecutor.WebRequest.KeepAlive = false;
                    _ctx.PendingRequest.RequestExecutor.WebRequest.Timeout = System.Threading.Timeout.Infinite;
                    _ctx.PendingRequest.RequestExecutor.WebRequest.ReadWriteTimeout = System.Threading.Timeout.Infinite;
                    ServicePointManager.DefaultConnectionLimit = 200;
                    ServicePointManager.MaxServicePointIdleTime = 2000;
                    ServicePointManager.MaxServicePoints = 1000;
                    ServicePointManager.SetTcpKeepAlive(false, 0, 0);
                    ServicePointManager.DnsRefreshTimeout = 2000;
                }
                else
                {
                    throw new ArgumentException(string.Format("{0} is null or empty", siteURL), "SiteCollection");
                }
            }
        }

        public SPLibraryService SPLibraryService
        {
            get
            {
                if (_ctx == null)
                {
                    _ctx = new ClientContext(_url)
                    {
                        Credentials = _credencial,
                        RequestTimeout = System.Threading.Timeout.Infinite
                    };
                    _ctx.PendingRequest.RequestExecutor.RequestKeepAlive = false;
                    _ctx.PendingRequest.RequestExecutor.WebRequest.KeepAlive = false;
                    _ctx.PendingRequest.RequestExecutor.WebRequest.Timeout = System.Threading.Timeout.Infinite;
                    _ctx.PendingRequest.RequestExecutor.WebRequest.ReadWriteTimeout = System.Threading.Timeout.Infinite;
                    ServicePointManager.DefaultConnectionLimit = 200;
                    ServicePointManager.MaxServicePointIdleTime = 2000;
                    ServicePointManager.MaxServicePoints = 1000;
                    ServicePointManager.SetTcpKeepAlive(false, 0, 0);
                    ServicePointManager.DnsRefreshTimeout = 2000;
                }

                if (spls == null || _ctx == null)
                {
                    spls = new SPLibraryService(_ctx);
                }
                return spls;
            }
        }

        public SPListService SPListService
        {
            get
            {
                if (_ctx == null)
                {
                    _ctx = new ClientContext(_url);
                    _ctx.Credentials = _credencial;
                    _ctx.RequestTimeout = System.Threading.Timeout.Infinite;
                    _ctx.PendingRequest.RequestExecutor.RequestKeepAlive = false;
                    _ctx.PendingRequest.RequestExecutor.WebRequest.KeepAlive = false;
                    _ctx.PendingRequest.RequestExecutor.WebRequest.Timeout = System.Threading.Timeout.Infinite;
                    _ctx.PendingRequest.RequestExecutor.WebRequest.ReadWriteTimeout = System.Threading.Timeout.Infinite;
                    ServicePointManager.DefaultConnectionLimit = 200;
                    ServicePointManager.MaxServicePointIdleTime = 2000;
                    ServicePointManager.MaxServicePoints = 1000;
                    ServicePointManager.SetTcpKeepAlive(false, 0, 0);
                    ServicePointManager.DnsRefreshTimeout = 2000;
                }

                if (spList == null || _ctx == null)
                {
                    spList = new SPListService(_ctx);
                }
                return spList;

            }
        }
        public SPGroupService SPGroupService
        {
            get
            {
                if (_ctx == null)
                {
                    _ctx = new ClientContext(_url);
                    _ctx.Credentials = _credencial;
                    _ctx.RequestTimeout = System.Threading.Timeout.Infinite;
                    _ctx.PendingRequest.RequestExecutor.RequestKeepAlive = false;
                    _ctx.PendingRequest.RequestExecutor.WebRequest.KeepAlive = false;
                    _ctx.PendingRequest.RequestExecutor.WebRequest.Timeout = System.Threading.Timeout.Infinite;
                    _ctx.PendingRequest.RequestExecutor.WebRequest.ReadWriteTimeout = System.Threading.Timeout.Infinite;
                    ServicePointManager.DefaultConnectionLimit = 200;
                    ServicePointManager.MaxServicePointIdleTime = 2000;
                    ServicePointManager.MaxServicePoints = 1000;
                    ServicePointManager.SetTcpKeepAlive(false, 0, 0);
                    ServicePointManager.DnsRefreshTimeout = 2000;
                }

                if (spGroups == null || _ctx == null)
                {
                    spGroups = new SPGroupService(_ctx);
                }
                return spGroups;
            }
        }

        public SPUserService SPUserService
        {
            get
            {
                if (_ctx == null)
                {
                    _ctx = new ClientContext(_url);
                    _ctx.Credentials = _credencial;
                    _ctx.RequestTimeout = System.Threading.Timeout.Infinite;
                    _ctx.PendingRequest.RequestExecutor.RequestKeepAlive = false;
                    _ctx.PendingRequest.RequestExecutor.WebRequest.KeepAlive = false;
                    _ctx.PendingRequest.RequestExecutor.WebRequest.Timeout = System.Threading.Timeout.Infinite;
                    _ctx.PendingRequest.RequestExecutor.WebRequest.ReadWriteTimeout = System.Threading.Timeout.Infinite;
                    ServicePointManager.DefaultConnectionLimit = 200;
                    ServicePointManager.MaxServicePointIdleTime = 2000;
                    ServicePointManager.MaxServicePoints = 1000;
                    ServicePointManager.SetTcpKeepAlive(false, 0, 0);
                    ServicePointManager.DnsRefreshTimeout = 2000;
                }
                if (spUsers == null)
                {
                    spUsers = new SPUserService(_ctx);
                }
                return spUsers;
            }
        }

        public IEnumerable<string> GetAllLibrarysName()
        {
            try
            {
                var oListsName = new List<string>();
                if (_ctx == null)
                {
                    _ctx = new ClientContext(_url);
                    _ctx.Credentials = _credencial;
                    _ctx.RequestTimeout = System.Threading.Timeout.Infinite;
                    _ctx.PendingRequest.RequestExecutor.RequestKeepAlive = false;
                    _ctx.PendingRequest.RequestExecutor.WebRequest.KeepAlive = false;
                    _ctx.PendingRequest.RequestExecutor.WebRequest.Timeout = System.Threading.Timeout.Infinite;
                    _ctx.PendingRequest.RequestExecutor.WebRequest.ReadWriteTimeout = System.Threading.Timeout.Infinite;
                    ServicePointManager.DefaultConnectionLimit = 200;
                    ServicePointManager.MaxServicePointIdleTime = 2000;
                    ServicePointManager.MaxServicePoints = 1000;
                    ServicePointManager.SetTcpKeepAlive(false, 0, 0);
                    ServicePointManager.DnsRefreshTimeout = 2000;
                }

                var oWeb = _ctx.Web;
                var oLists = oWeb.Lists;
                var query = from oList in oLists
                            where oList.BaseType == BaseType.DocumentLibrary
                            select oList;

                var resultList = _ctx.LoadQuery(query);
                _ctx.ExecuteQuery();

                foreach (var list in resultList)
                {
                    oListsName.Add(list.Title);
                }

                return oListsName;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposed)
                return;

            if (disposing)
            {
                _url = string.Empty;
                spGroups = null;
                spList = null;
                spls = null;
                spUsers = null;

                _ctx.Dispose();
                _ctx = null;
                _credencial = null;
            }

            disposed = true;
        }
    }
}
