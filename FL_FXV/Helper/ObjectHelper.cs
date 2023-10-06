using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Columns;
using DevExpress.XtraGrid.Views.Grid;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FL_FXV.Helper
{
    public class ObjectHelper
    {
        public static object inJect(object source, object coppy)
        {
            Type t = source.GetType();
            PropertyInfo[] props = t.GetProperties();
            Type typeCopy = coppy.GetType();
            PropertyInfo[] c = typeCopy.GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (var prop in props)
            {
                foreach (var item in c)
                {
                    if (prop.Name == item.Name)
                    {
                        item.SetValue(coppy, prop.GetValue(source), null);
                    }
                }
            }
            return coppy;
        }

        public static void CommitGridToDataSource(GridControl grcData)
        {
            if (grcData != null && grcData.MainView != null)
            {
                if (grcData.MainView.PostEditor() && grcData.MainView.UpdateCurrentRow())
                {
                    grcData.MainView.CloseEditor();
                }
            }
        }
    }
}
