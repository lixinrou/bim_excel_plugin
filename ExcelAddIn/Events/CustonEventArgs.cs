﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAddIn.Events
{
    public class CustonEventArgs:EventArgs
    {
        public string Text { get; set; }
        public string ObjectId { get; set; }
        public int Count { get; set; }
    }
}
