using TakeItEasy.DatabaseSrc;

namespace TakeItEasy
{

    using System.Collections.Generic;
    using System.IO;
    using WpfControls.Editors;
    using System.Linq;
    using System.Threading;
    using System.Collections.ObjectModel;
    using System;

    public class DbSuggestionProvider : ISuggestionProvider
    {

        public System.Collections.IEnumerable GetSuggestions(string filter)
        {
            try
            {
                if (string.IsNullOrEmpty(filter))
                {
                    return null;
                }

                ObservableCollection<ObjectData> lst = new ObservableCollection<ObjectData>();
                string[] parseItems = filter.Split('.');
                //database level
                if (parseItems.Length <= 1)
                {
                    lst = GetDBAction.GetDBList();
                }
                //fields level
                else if (parseItems.Length == 4)
                {
                    lst = GetDBAction.GetFieldList(parseItems[0], parseItems[2]);
                }
                //table level
                else if (parseItems.Length == 3)
                {
                    lst = GetDBAction.GetTableList(new ObjectData(parseItems[0], 0));
                }
                else
                {
                    lst = null;
                }
                //get suggestion list
                if (lst != null)
                {
                    List<ObjectData> lst1 = new List<ObjectData>();
                    foreach (ObjectData obj in lst)
                    {
                        if (obj.Name.StartsWith(parseItems[parseItems.Length - 1], System.StringComparison.InvariantCultureIgnoreCase))
                        {
                            lst1.Add(obj);
                        }
                    }
                    return lst1;
                }

                return null;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

    }

}
