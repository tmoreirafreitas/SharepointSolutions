using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace SPUtilitieServicesSolution
{
    public class SPGroupService
    {
        public IList<SPUserService> UserList { get; private set; }
        private readonly ClientContext _ctx;
        public SPGroupService(ClientContext ctx)
        {
            _ctx = ctx;
            UserList = new List<SPUserService>();
        }

        public IEnumerable<string> GetAllGroupsName()
        {
            try
            {
                var groupsName = new List<string>();
                var groups = _ctx.Web.SiteGroups;
                _ctx.Load(groups, grp => grp.Include(group => group.Title, group => group.LoginName));
                _ctx.ExecuteQuery();
                foreach (var group in groups)
                {
                    groupsName.Add(group.Title);
                }

                return groupsName;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public DataTable GetAllUsersByGroup(string groupName)
        {
            try
            {
                var objs = new List<IDictionary<string, string>>();
                var groups = _ctx.Web.SiteGroups;
                var queryGroup = ((from oGroup in groups
                                   where oGroup.Title == groupName
                                   select oGroup)
                                   .Include(g => g.Users.Include(u => u.LoginName, u => u.Title, u => u.Email, u => u.Id), g => g.Title));
                var result = _ctx.LoadQuery(queryGroup);
                _ctx.Load(_ctx.Web, w => w.Title);
                _ctx.ExecuteQuery();

                var group = result.First();
                foreach (var user in group.Users)
                {
                    var DyObjectsList = new Dictionary<string, string>();
                    DyObjectsList.Add("Title", user.Title);
                    DyObjectsList.Add("LoginName", user.LoginName);
                    DyObjectsList.Add("E-mail", user.Email);
                    DyObjectsList.Add("GroupName", group.Title);
                    DyObjectsList.Add("Campanha", _ctx.Web.Title);
                    DyObjectsList.Add("Url", _ctx.Url);
                    objs.Add(DyObjectsList);
                }

                return Util.GetDataTableFromListDictionary(objs);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void Create(string title, string description, string strOwner, string[] permissions = null, IList<string> users = null)
        {
            try
            {
                if (string.IsNullOrEmpty(title))
                {
                    throw new ArgumentException("O title precisa ser informado");
                }
                if (string.IsNullOrEmpty(strOwner))
                {
                    throw new ArgumentException("O Owner precisa ser informado");
                }
                if (permissions == null || permissions.Length == 0)
                {
                    throw new ArgumentException("É necessário informar ao menos uma permissão para o grupo");
                }

                GroupCreationInformation groupCreationInfo = new GroupCreationInformation
                {
                    Title = title,
                    Description = description
                };

                User owner = _ctx.Web.EnsureUser(strOwner);
                Group group = _ctx.Web.SiteGroups.Add(groupCreationInfo);
                RoleDefinitionBindingCollection objBindingColl = new RoleDefinitionBindingCollection(_ctx);
                if (permissions != null && permissions.Length > 0)
                {
                    foreach (var permission in permissions)
                    {
                        RoleDefinition customPermissionRole = _ctx.Web.RoleDefinitions.GetByName(permission);
                        objBindingColl.Add(customPermissionRole);
                    }
                }
                _ctx.Web.RoleAssignments.Add(group, objBindingColl);
                if (users != null)
                {
                    foreach (var user in users)
                    {
                        User member = _ctx.Web.EnsureUser(user);
                        group.Users.AddUser(member);
                    }
                }

                group.Update();
                _ctx.ExecuteQuery();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void DeleteGroup(string groupName)
        {
            if (string.IsNullOrEmpty(groupName))
            {
                throw new ArgumentException(string.Format("{0} não pode ser vazio e nem nulo", groupName), "groupName");
            }

            try
            {
                var groups = _ctx.Web.SiteGroups;
                _ctx.Load(groups, grp => grp.Include(group => group.Title, group => group.LoginName)
                .Where(g => g.Title == groupName));
                _ctx.ExecuteQuery();
                if (groups != null && groups.Count > 0)
                    groups.Remove(groups[0]);
                _ctx.ExecuteQuery();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void AddUserToGroup(string userTitle, string groupName)
        {
            if (string.IsNullOrEmpty(groupName))
            {
                throw new ArgumentException(string.Format("{0} não pode ser vazio e nem nulo", groupName), "groupName");
            }
            if (string.IsNullOrEmpty(userTitle))
            {
                throw new ArgumentException(string.Format("{0} não pode ser vazio e nem nulo", userTitle), "userTitle");
            }

            try
            {
                var groups = _ctx.Web.SiteGroups;
                var queryGroup = ((from oGroup in groups
                                   where oGroup.Title == groupName
                                   select oGroup)
                                   .Include(g => g.Users.Include(u => u.LoginName, u => u.Title, u => u.Email).Where(u => u.Title != userTitle)));
                var result = _ctx.LoadQuery(queryGroup);
                _ctx.ExecuteQuery();

                if (result != null && result.Count() > 0)
                {
                    var group = result.First();
                    var user = _ctx.Web.EnsureUser(userTitle);
                    group.Users.AddUser(user);
                    group.Update();
                    _ctx.ExecuteQuery();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void AddMultipleUserToGroup(IEnumerable<string> listUserTitle, string groupName, IProgress<string> progress = null)
        {
            if (string.IsNullOrEmpty(groupName))
            {
                throw new ArgumentException(string.Format("{0} não pode ser vazio e nem nulo", groupName), "groupName");
            }
            if (listUserTitle == null || listUserTitle.Count() == 0)
            {
                throw new ArgumentException(string.Format("{0} não pode ser vazio e nem nulo", groupName), "listUserTitle");
            }

            try
            {               
                var groups = _ctx.Web.SiteGroups;
                var queryGroup = ((from oGroup in groups
                                   where oGroup.Title == groupName
                                   select oGroup)
                                   .Include(g => g.Users.Include(u => u.LoginName, u => u.Title, u => u.Email, u => u.Id)));
                var result = _ctx.LoadQuery(queryGroup);
                _ctx.ExecuteQuery();

                if (result != null && result.Count() > 0)
                {
                    var group = result.First();
                    _ctx.Load(group);
                    _ctx.ExecuteQuery();
                    
                    foreach (var userTitle in listUserTitle)
                    {
                        try
                        {
                            var user = _ctx.Web.EnsureUser(userTitle);
                            group.Users.AddUser(user);
                            group.Update();

                            if (progress != null)
                            {
                                string message = string.Format("Site:[{0}]:\n\nUsuário: [{1}] foi adicionado no grupo: [{2}]", _ctx.Url, userTitle, group.Title);
                                progress.Report(message);
                            }
                        }
                        catch (Exception ex)
                        {
                            if (progress != null)
                            {
                                string message = string.Format("Site:[{0}]:\n\nUsuário: [{1}] não pode ser inserido no grupo: [{2}]\n\nMensage: {3}\n\nStackTrace: {4}", _ctx.Url, userTitle, group.Title, ex.Message, ex.StackTrace);
                                progress.Report(message);
                            }
                        }
                    }
                    _ctx.ExecuteQuery();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
