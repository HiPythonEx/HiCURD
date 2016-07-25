
        private bool OnDelete(DataGridView dgvMain, int index, #table#VO vo)
        {
            if (vo == null)
            {
                return false;
            }
            //string id = DGViewUtil.GetCellValue(dgvMain, index, 1);
            string name = DGViewUtil.GetCellValue(dgvMain, index, 2);
            bool isOK = MsgBoxHelper.Confirm(string.Format("您确定要删除名称为[{0}]的#cname#?", name));
            if (!isOK)
            {
                return false;
            }
            //vo.TypeID = id;
			// your implement
            throw new NotImplementedException("this function not support in rest");
        }


