using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SPUtilitieServicesSolution.Extensions
{
    public static class CSOMExtensions
    {
        public static Task ExecuteQueryAsync(this ClientContext clientContext)
        {
            return Task.Factory.StartNew(() =>
            {
                clientContext.ExecuteQuery();
            });
        }
    }
}
