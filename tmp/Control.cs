using System;
using System.Windows.Forms;

using HiCSUIHelper;
using FormTest.Logic;

using Model;

namespace FormTest
{
    sealed class Controls
    {

        public void OnAction(string id)
        {
            if (string.IsNullOrWhiteSpace(id))
            {
                return;
            }
			
            HiManageBase form = CreateForm(id);
            if (form == null)
            {
                return;
            }
            form.OnDelete = OnDelete;
            form.OnEdit = OnEdit;
            form.Show();
        }

        private bool OnEdit(DataGridView dgvMain, int index, object vo, HiManageBase logic)
        {
            switch (logic.LogicInfo.Tag)
            {
 #onedit#
                default:
                    {
                        return false;
                    }
            }
        }

        private bool OnDelete(DataGridView dgvMain, int index, object vo, HiManageBase logic)
        {
            switch (logic.LogicInfo.Tag)
            {
 #ondelete#
                default:
                    {
                        return false;
                    }
            }
        }
        private HiManageBase CreateForm(string id)
        {
            switch (id)
            {
 #create_form#
                default:
                    {
                        return null;
                    }
            }
        }
		
#edit_funs#

#del_funs#

        private HiManageBase CreateForm<T>(string formId) where T : class, new()
        {
            string json = ViewConfig.GetView(formId + ".Form");
            if (string.IsNullOrWhiteSpace(json))
            {
                return null;
            }
            return new ProductTypeMng<T>(HiCSUtil.Json.Json2Obj<LogicInfo>(json));
        }

        private bool OnDelete(DataGridView dgvMain, int index, object vo)
        {
            return false;
        }

        private bool OnEdit(DataGridView dgvMain, int index, object vo)
        {
            return false;
        }

    }
}
