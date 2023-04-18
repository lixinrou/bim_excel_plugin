using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAddIn.Data
{
    public class ExcuteSql
    {
        public static string SYSTEM_FILTER = @"SELECT Id  FROM `bim_t_element` WHERE `RowId` IN (SELECT Element_RowId  FROM `bim_t_items` 
                                               WHERE `Key` = '系统名称' AND `Value` IN('{0}')) AND (ModelFileId = '{1}' OR Id LIKE '%{1}%') ORDER BY CreateTime DESC;";
    }
}
