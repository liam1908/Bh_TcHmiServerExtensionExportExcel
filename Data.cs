//-----------------------------------------------------------------------
// <copyright file="Data.cs" company="Beckhoff Automation GmbH & Co. KG">
//     Copyright (c) Beckhoff Automation GmbH & Co. KG. All Rights Reserved.
// </copyright>
//-----------------------------------------------------------------------

namespace SE_Test
{
    // Contains runtime data. This class should be thread-safe.
    public class Data
    {
        private int maxRandom = 1000;
        private object maxRandomLock = new object();

        private bool boolTestVar = false;
        private object boolTestVarLock = new object();

        private bool exportExcel = false;
        private object exportExcelLock = new object();

        

        public int MaxRandom
        {
            get
            {
                lock (this.maxRandomLock)
                {
                    return this.maxRandom;
                }
            }

            set
            {
                lock (this.maxRandomLock)
                {
                    this.maxRandom = value;
                }
            }
        }

        public bool BoolTestVar
        {
            get
            {
                lock (this.boolTestVarLock)
                {
                    return this.boolTestVar;
                }
            }

            set
            {
                lock (this.boolTestVarLock)
                {
                    this.boolTestVar = value;
                }
            }
        }

        public bool ExportExcel
        {
            get
            {
                lock (this.exportExcelLock)
                {
                    return this.exportExcel;
                }
            }

            set
            {
                lock (this.exportExcelLock)
                {
                    this.exportExcel = value;
                }
            }
        }
    }
}
