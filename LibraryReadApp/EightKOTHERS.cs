using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace LibraryReadApp
{
    public class EightKOTHERS
    {
        public static DataTable dtResponse = new DataTable();

        public EightKOTHERS()
        {
            dtResponse.Columns.Add(Constants.LibraryName);
            dtResponse.Columns.Add(Constants.R_object_type);
            dtResponse.Columns.Add(Constants.Documentum_i_chronicle_id);
            dtResponse.Columns.Add(Constants.Documentum_r_object_id);
            dtResponse.Columns.Add(Constants.Documentum_content_id);
            dtResponse.Columns.Add(Constants.Documentum_r_folder_path);
            dtResponse.Columns.Add(Constants.Title);
            dtResponse.Columns.Add(Constants.I_full_format);
            dtResponse.Columns.Add(Constants.A_webc_url);
            dtResponse.Columns.Add(Constants.R_object_id);
            dtResponse.Columns.Add(Constants.I_chronicle_id);
            dtResponse.Columns.Add(Constants.Content_id);
            dtResponse.Columns.Add(Constants.R_folder_path);
            dtResponse.Columns.Add(Constants.Display_Order);
        }

        public DataTable GetListFileId()
        {
            SecureString secureString = null;
            List list = null;
            ListItemCollection items = null;
            string Name = string.Empty, guid = string.Empty, Version = string.Empty, display_order = string.Empty,
                documentum_i_chronicle_id = string.Empty, documentum_r_object_id = string.Empty, documentum_content_id = string.Empty;
            try
            {
                using (ClientContext clientContext = new ClientContext(ConfigurationManager.AppSettings.Get(SPOConstants.SPOSiteURL)))
                {
                    secureString = new NetworkCredential("", ConfigurationManager.AppSettings.Get(SPOConstants.SPOPassword)).SecurePassword;
                    clientContext.Credentials = new SharePointOnlineCredentials(ConfigurationManager.AppSettings.Get(SPOConstants.SPOUserName), secureString);
                    list = clientContext.Web.Lists.GetByTitle(ConfigurationManager.AppSettings.Get(SPOConstants.SPOFolder8KOther));

                    CamlQuery query = CamlQuery.CreateAllItemsQuery();
                    items = list.GetItems(query);
                    clientContext.Load(items);
                    clientContext.ExecuteQuery();

                    foreach (var obj in items)
                    {
                        ListItem listItem = obj.File.ListItemAllFields;
                        clientContext.Load(listItem, item => item[SPOConstants.Name],
                                                 item => item[SPOConstants.UniqueId],
                                                 item => item[SPOConstants.UIVersionString],
                                                 item => item[SPOConstants.SortOrder],
                                                 //item => item[SPOConstants.Documentum_i_chronicle_id],
                                                 item => item[SPOConstants.Documentum_r_object_id]);
                        //item => item[SPOConstants.Documentum_content_id]);

                        clientContext.ExecuteQuery();

                        Dictionary<string, object> keyValuePairs = obj.FieldValues;

                        Name = keyValuePairs.ContainsKey(SPOConstants.Name) ? (keyValuePairs[SPOConstants.Name] != null ? keyValuePairs[SPOConstants.Name].ToString() : "") : "";
                        guid = keyValuePairs.ContainsKey(SPOConstants.UniqueId) ? (keyValuePairs[SPOConstants.UniqueId] != null ? keyValuePairs[SPOConstants.UniqueId].ToString() : "") : "";
                        Version = keyValuePairs.ContainsKey(SPOConstants.UIVersionString) ? (keyValuePairs[SPOConstants.UIVersionString] != null ? keyValuePairs[SPOConstants.UIVersionString].ToString() : "") : "";
                        documentum_i_chronicle_id = keyValuePairs.ContainsKey(SPOConstants.Documentum_i_chronicle_id) ? (keyValuePairs[SPOConstants.Documentum_i_chronicle_id] != null ? keyValuePairs[SPOConstants.Documentum_i_chronicle_id].ToString() : "") : "";
                        documentum_r_object_id = keyValuePairs.ContainsKey(SPOConstants.Documentum_r_object_id) ? (keyValuePairs[SPOConstants.Documentum_r_object_id] != null ? keyValuePairs[SPOConstants.Documentum_r_object_id].ToString() : "") : "";
                        documentum_content_id = keyValuePairs.ContainsKey(SPOConstants.Documentum_content_id) ? (keyValuePairs[SPOConstants.Documentum_content_id] != null ? keyValuePairs[SPOConstants.Documentum_content_id].ToString() : "") : "";
                        display_order = keyValuePairs.ContainsKey(SPOConstants.SortOrder) ? (keyValuePairs[SPOConstants.SortOrder] != null ? keyValuePairs[SPOConstants.SortOrder].ToString() : "0") : "0";

                        dtResponse.Rows.Add(ConfigurationManager.AppSettings.Get(SPOConstants.SPOFolder8KOther), Constants.Ir_8k_filing, documentum_r_object_id, documentum_i_chronicle_id, documentum_content_id, Constants.EightK_Other_Filing_r_folder_path, Name, FileFormatConstants.PDF, Constants.EightK_Other_Filing + "/" + Name, guid, guid, guid + '@' + guid, Constants.EightK_Other_Filing, display_order);

                    }
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine("Error in 8K&Others :" + ex.Message);
                dtResponse = null;
            }
            return dtResponse;
        }
    }
}
