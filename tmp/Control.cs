using System;

using BomManage.Logic;

using BomManage.Model.VO;

namespace BomManage.View.ControlPlus
{
    partial class Controls
    {

        public static object CreateObject(LogicInfo info)
        {
            if (info == null)
            {
                return null;
            }

            switch (info.Tag)
            {
#create_form#
                default:
                    {
                        return null;
                    }
            }
        }
    }
}