using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Data;
using System.Dynamic;
using System.Linq;
using System.Net;
using System.Text;

namespace SPUtilitieServicesSolution
{
    public class SPUserService
    {
        private readonly ClientContext _ctx;

        public SPUserService(ClientContext ctx)
        {
            _ctx = ctx;
        }

        public DataTable GetAllUsers()
        {
            var objs = new List<IDictionary<string, object>>();
            try
            {
                var list = _ctx.Web.SiteUserInfoList;
                var users = list.GetItems(CamlQuery.CreateAllItemsQuery());
                _ctx.Load(list.Fields, fields => fields
                .Include(f => f.CanBeDeleted, f => f.InternalName, f => f.StaticName,
                f => f.FieldTypeKind, f => f.Title, f => f.Filterable, f => f.Sortable)
                .Where(f => f.FieldTypeKind != FieldType.Attachments && f.CanBeDeleted == true));
                _ctx.Load(users);
                _ctx.ExecuteQuery();

                foreach (var user in users)
                {
                    _ctx.Load(user);
                    _ctx.ExecuteQuery();
                    dynamic spuser = new ExpandoObject();

                    foreach (var key in user.FieldValues.Keys)
                    {
                        if (user.FieldValues[key] != null)
                        {
                            if (user.FieldValues[key].GetType() == typeof(FieldUserValue))
                            {
                                Object objFuv = null;
                                user.FieldValues.TryGetValue(key, out objFuv);
                                if (objFuv != null)
                                {
                                    var fuv = user.FieldValues[key] as FieldUserValue[];
                                    if (fuv != null)
                                    {
                                        for (int i = 0; i < fuv.Length; i++)
                                        {
                                            dynamic oUser = new ExpandoObject();
                                            oUser.Id = fuv[i].LookupId;
                                            oUser.Name = fuv[i].LookupValue;
                                            Util.AddProperty(spuser, string.Format("{0}{1}", key, oUser.Id), oUser.Id + ";#" + oUser.Name);
                                        }
                                    }
                                    else
                                    {
                                        Object objUser = null;
                                        user.FieldValues.TryGetValue(key, out objUser);
                                        if (objUser != null)
                                        {
                                            dynamic oUser = new ExpandoObject();
                                            oUser.Id = ((FieldUserValue)user.FieldValues[key]).LookupId;
                                            oUser.Name = ((FieldUserValue)user.FieldValues[key]).LookupValue;
                                            Util.AddProperty(spuser, key, oUser.Id + ";#" + oUser.Name);
                                        }
                                    }
                                }
                            }
                            else if (user.FieldValues[key].GetType() == typeof(FieldLookupValue))
                            {
                                Object objFl = null;
                                user.FieldValues.TryGetValue(key, out objFl);
                                if (objFl != null)
                                {
                                    var flv = user.FieldValues[key] as FieldLookupValue[];
                                    if (flv != null)
                                    {
                                        for (int i = 0; i < flv.Length; i++)
                                        {
                                            dynamic fl = new ExpandoObject();
                                            fl.Id = flv[i].LookupId;
                                            fl.Name = flv[i].LookupValue;
                                            Util.AddProperty(spuser, string.Format("{0}{1}", key, fl.Id), fl);
                                        }
                                    }
                                    else
                                    {
                                        Object obj = null;
                                        user.FieldValues.TryGetValue(key, out obj);
                                        if (obj != null)
                                        {
                                            dynamic fl = new ExpandoObject();
                                            fl.Id = ((FieldLookupValue)user.FieldValues[key]).LookupId;
                                            fl.Name = ((FieldLookupValue)user.FieldValues[key]).LookupValue;
                                            Util.AddProperty(spuser, key, fl);
                                        }
                                    }
                                }
                            }
                            else
                            {
                                Object obj = null;
                                user.FieldValues.TryGetValue(key, out obj);
                                if (obj != null)
                                {
                                    if (user[key] != null)
                                        Util.AddProperty(spuser, key, user.FieldValues[key]);
                                    else
                                        Util.AddProperty(spuser, key, "");
                                }
                            }
                        }
                        else
                        {
                            Util.AddProperty(spuser, key, "");
                        }
                    }
                    objs.Add(spuser);
                }
                return Util.GetDataTableFromListDynamic(objs);
            }

            catch (Exception ex)
            {
                throw ex;
            }
        }

        public bool DeleteUserById(int id)
        {
            var isDeleted = false;

            if (id == 0)
            {
                throw new ArgumentException(string.Format("{0} can't be zero.", id), "id");
            }
            try
            {
                var oList = _ctx.Web.SiteUserInfoList;
                _ctx.Load(oList);
                _ctx.Load(oList.Fields, fields => fields
                .Include(f => f.CanBeDeleted,
                f => f.InternalName, f => f.StaticName, f => f.FieldTypeKind, f => f.Title)
                .Where(f => f.FieldTypeKind != FieldType.Attachments && f.CanBeDeleted));
                _ctx.ExecuteQuery();
                if (oList != null && oList.ItemCount > 0)
                {
                    var oListItem = oList.GetItemById(id);
                    oListItem.DeleteObject();
                    _ctx.ExecuteQuery();
                    isDeleted = true;
                }
                else
                {
                    isDeleted = false;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return isDeleted;
        }
    }
}