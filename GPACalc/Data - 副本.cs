using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace GPACalc
{
    class Data
    {
        /// <summary>
        /// Generate all the custom datatable columns
        /// </summary>
        /// <param name="dt">custome datatable</param>
        public static void GenerateCustomeDataTableColumns(ref DataTable dt)
        {
            dt.Columns.Add("Student ID");
            dt.Columns.Add("Student Name");
            dt.Columns.Add("Course Name");
            dt.Columns.Add("Credits");
            dt.Columns.Add("Score");
        }

        /// <summary>
        /// Generate all the Calculated datatable columns
        /// </summary>
        /// <param name="dt">Calculated datatable</param>
        public static void GenerateCalculatedDataTableColumns(ref DataTable dt)
        {
            dt.Columns.Add("Student ID");
            dt.Columns.Add("Student Name");
            dt.Columns.Add("Course Name");
            dt.Columns.Add("Credits");
            dt.Columns.Add("Score");
            dt.Columns.Add("GPA");

            // Change the datacolunm type
            dt.Columns[0].DataType = Type.GetType("System.Int64");
            dt.Columns[3].DataType = Type.GetType("System.Single");
            dt.Columns[4].DataType = Type.GetType("System.Single");
            dt.Columns[5].DataType = Type.GetType("System.Single");
        }

        /// <summary>
        /// Generate all the user custom gpa methods dataset columns
        /// </summary>
        /// <param name="ds">custom dataset</param>
        public static void GenerateCustomGPAMethodsDataSetColumns(ref DataSet ds)
        {
            for (int i = 0; i < 3; i++)
            {
                ds.Tables.Add(i.ToString());
                ds.Tables[i.ToString()].Columns.Add("Score From");
                ds.Tables[i.ToString()].Columns[0].DataType = Type.GetType("System.Single");
                ds.Tables[i.ToString()].Columns.Add("GPA");
                ds.Tables[i.ToString()].Columns[1].DataType = Type.GetType("System.Single");
            }
        }

        /// <summary>
        /// Generate all the gpa methods dataset columns
        /// </summary>
        /// <param name="ds">dataset</param>
        public static void GenerateGPAMethodsDataSetColumns(ref DataSet ds)
        {
            for (int i = 0; i < 7; i++)
            {
                ds.Tables.Add(i.ToString());
                ds.Tables[i.ToString()].Columns.Add("Score From");
                ds.Tables[i.ToString()].Columns[0].DataType = Type.GetType("System.Single");
                ds.Tables[i.ToString()].Columns.Add("GPA");
                ds.Tables[i.ToString()].Columns[1].DataType = Type.GetType("System.Single");
            }
        }

        /// <summary>
        /// calculate the gpas
        /// </summary>
        /// <param name="inDataTable">in datatable that contains all the data</param>
        /// <param name="outDataTable">out datatable that contains seperate gpa</param>
        /// <param name="outOverallDataTable">out datatable that contains overall gpa</param>
        /// <param name="dtMethod">the method to calculate with</param>
        public static void CalculteGPAs(ref DataTable inDataTable, ref DataTable outDataTable, ref DataTable outOverallDataTable, DataTable dtMethod)
        {
            // Store the id of the student to compare with
            long sid = long.Parse(inDataTable.Rows[0][0].ToString());

            // Store the gpas that the student already earned
            float gpas = 0;
            // and the credits that the student already have
            float credits = 0;

            // sort the student id column for it is the rule when calculate
            inDataTable.DefaultView.Sort = "Student ID";

            //////////////////////////////////////////////////////////////////////////
            System.Diagnostics.Stopwatch sw = new System.Diagnostics.Stopwatch();
           
            foreach (DataRow dr in inDataTable.Rows)
            {
                if (long.Parse(dr["Student ID"].ToString()) == sid)
                {

                    // Create a new datarow contains all the information of dr and calculate tht gpa
                    DataRow drn = outDataTable.NewRow();

                    // change the string gradt to number grades
                    switch (dr[4].ToString())
                    {
                        case "优秀": dr[4] = 90; break;
                        case "良好": dr[4] = 80; break;
                        case "中等": dr[4] = 70; break;
                        case "合格": dr[4] = 60; break;
                        case "不合格": dr[4] = 0; break;

                        case "及格": dr[4] = 60; break;
                        case "不及格": dr[4] = 0; break;
                        case "": dr[4] = 0; break;
                    }

                    // copy all the data from orginal dr to new drn
                    for (int i = 0; i < 5; i++)
                    {
                        drn[i] = dr[i];
                    }


                    // compare and calculate gpa
                    for (int i = 0; i < dtMethod.Rows.Count; i++)
                    {
                        if (float.Parse(drn[4].ToString()) >= float.Parse(dtMethod.Rows[i][0].ToString()))
                        {
                            drn[5] = dtMethod.Rows[i][1];
                            break;
                        }
                    }

                    sw.Start();
                    // add to the datatable
                    outDataTable.Rows.Add(drn);
                    sw.Stop();
                    gpas += float.Parse(drn[3].ToString()) * float.Parse(drn[5].ToString());
                    credits += float.Parse(drn[3].ToString());
                }
                else
                {
                    // create a new datarow to store the overall gpa
                    DataRow drn = outOverallDataTable.NewRow();

                    // copy all the data from original dr to new drn (notice that those info are not in current dr, they are in the last one)
                    for (int i = 0; i < 2; i++)
                    {
                        drn[i] = outDataTable.Rows[outDataTable.Rows.Count - 1][i];
                    }

                    // calculate the overall gpa
                    drn[2] = gpas / credits;

                    // add it to the datatable
                    outOverallDataTable.Rows.Add(drn);

                    // reset the ints
                    gpas = credits = 0;

                    //////////////////////////////////////////////////////////////////////////
                    // do all those calculate for a whole new person
                    sid = long.Parse(dr[0].ToString());

                    // Create a new datarow contains all the information of dr and calculate tht gpa
                    DataRow drnn = outDataTable.NewRow();

                    // change the string gradt to number grades
                    switch (dr[4].ToString())
                    {
                        case "优秀": dr[4] = 90; break;
                        case "良好": dr[4] = 80; break;
                        case "中等": dr[4] = 70; break;
                        case "合格": dr[4] = 60; break;
                        case "不合格": dr[4] = 0; break;

                        case "及格": drn[4] = 60; break;
                        case "不及格": drn[4] = 0; break;
                        case "": dr[4] = 0; break;
                    }

                    // copy all the data from orginal dr to new drn
                    for (int i = 0; i < 5; i++)
                    {
                        drnn[i] = dr[i];
                    }

                    // compare and calculate gpa
                    for (int i = 0; i < dtMethod.Rows.Count; i++)
                    {
                        if (float.Parse(drnn[4].ToString()) >= float.Parse(dtMethod.Rows[i][0].ToString()))
                        {
                            drnn[5] = dtMethod.Rows[i][1];
                            break;
                        }
                    }

                    // add to the datatable
                    outDataTable.Rows.Add(drnn);
                    gpas += float.Parse(drnn[3].ToString()) * float.Parse(drnn[5].ToString());
                    credits += float.Parse(drnn[3].ToString());
                }
            }

            //////////////////////////////////////////////////////////////////////////
            // calculate the last student's overall gpa

            // create a new datarow to store the overall gpa
            DataRow drnl = outOverallDataTable.NewRow();

            // copy all the data from original dr to new drn (notice that those info are not in current dr, they are in the last one)
            for (int i = 0; i < 2; i++)
            {
                drnl[i] = outDataTable.Rows[outDataTable.Rows.Count - 1][i];
            }

            // calculate the overall gpa
            drnl[2] = gpas / credits;

            // add it to the datatable
            outOverallDataTable.Rows.Add(drnl);


            System.Windows.Forms.MessageBox.Show(sw.ElapsedMilliseconds.ToString());
            //////////////////////////////////////////////////////////////////////////
        }

        /// <summary>
        /// generate all the gpa methods
        /// </summary>
        /// <param name="ds">dataset that contains all the methods</param>
        public static void GenerateCommonGPAMethods(ref DataSet ds)
        {
            DataRow dr;
 #region 百分算法
            DataTable dt = new DataTable() { TableName = "0" };

            // create the schema
            dt.Columns.Add("Score From");
            dt.Columns[0].DataType = Type.GetType("System.Single");
            dt.Columns.Add("GPA");
            dt.Columns[1].DataType = Type.GetType("System.Single");

            for (int i = 100; i >= 0; i--)
            {
                dr = dt.NewRow();
                dr[0] = dr[1] = i;
                dt.Rows.Add(dr);
            }

            // add to dataset
            ds.Tables.Add(dt);
#endregion

 #region 5分算法
            dt = new DataTable() { TableName = "1" };

            // create the schema
            dt.Columns.Add("Score From");
            dt.Columns[0].DataType = Type.GetType("System.Single");
            dt.Columns.Add("GPA");
            dt.Columns[1].DataType = Type.GetType("System.Single");

            dr = dt.NewRow();
            dr[0] = 100;
            dr[1] = 5.0;
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = 96;
            dr[1] = 4.8;
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = 93;
            dr[1] = 4.5;
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = 90;
            dr[1] = 4.0;
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = 86;
            dr[1] = 3.8;
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = 83;
            dr[1] = 3.5;
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = 80;
            dr[1] = 3.0;
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = 76;
            dr[1] = 2.8;
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = 73;
            dr[1] = 2.5;
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = 70;
            dr[1] = 2.0;
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = 66;
            dr[1] = 1.8;
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = 63;
            dr[1] = 1.5;
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = 60;
            dr[1] = 1.0;
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = 0;
            dr[1] = 0.0;
            dt.Rows.Add(dr);

            // add to dataset
            ds.Tables.Add(dt);
#endregion

#region 标准算法
            dt = new DataTable() { TableName = "2" };

            // create the schema
            dt.Columns.Add("Score From");
            dt.Columns[0].DataType = Type.GetType("System.Single");
            dt.Columns.Add("GPA");
            dt.Columns[1].DataType = Type.GetType("System.Single");

            dr = dt.NewRow();
            dr[0] = 90;
            dr[1] = 4.0;
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = 80;
            dr[1] = 3.0;
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = 70;
            dr[1] = 2.0;
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = 60;
            dr[1] = 1.0;
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = 0;
            dr[1] = 0;
            dt.Rows.Add(dr);

            // add to dataset
            ds.Tables.Add(dt);
#endregion

#region 北大算法
            dt = new DataTable() { TableName = "3" };

            // create the schema
            dt.Columns.Add("Score From");
            dt.Columns[0].DataType = Type.GetType("System.Single");
            dt.Columns.Add("GPA");
            dt.Columns[1].DataType = Type.GetType("System.Single");

            dr = dt.NewRow();
            dr[0] = 90;
            dr[1] = 4.0;
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = 85;
            dr[1] = 3.7;
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = 82;
            dr[1] = 3.3;
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = 78;
            dr[1] = 3.0;
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = 75;
            dr[1] = 2.7;
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = 72;
            dr[1] = 2.3;
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = 68;
            dr[1] = 2.0;
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = 64;
            dr[1] = 1.5;
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = 60;
            dr[1] = 1.0;
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = 0;
            dr[1] = 0;
            dt.Rows.Add(dr);

            // add to dataset
            ds.Tables.Add(dt);
#endregion

#region 浙大算法
            dt = new DataTable() { TableName = "4" };

            // create the schema
            dt.Columns.Add("Score From");
            dt.Columns[0].DataType = Type.GetType("System.Single");
            dt.Columns.Add("GPA");
            dt.Columns[1].DataType = Type.GetType("System.Single");

            dr = dt.NewRow();
            dr[0] = 85;
            dr[1] = 4.0;
            dt.Rows.Add(dr);

            for (int i = 24; i >= 0;i-- )
            {
                dr = dt.NewRow();
                dr[0] = 60 + i;
                dr[1] = 1.5 + 1.0 * i;
                dt.Rows.Add(dr);
            }

            dr = dt.NewRow();
            dr[0] = 0;
            dr[1] = 0;
            dt.Rows.Add(dr);

            // add to dataset
            ds.Tables.Add(dt);
#endregion

#region 上交算法
            dt = new DataTable() { TableName = "5" };

            // create the schema
            dt.Columns.Add("Score From");
            dt.Columns[0].DataType = Type.GetType("System.Single");
            dt.Columns.Add("GPA");
            dt.Columns[1].DataType = Type.GetType("System.Single");

            dr = dt.NewRow();
            dr[0] = 95;
            dr[1] = 4.3;
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = 90;
            dr[1] = 4.0;
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = 85;
            dr[1] = 3.7;
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = 80;
            dr[1] = 3.3;
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = 75;
            dr[1] = 3.0;
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = 70;
            dr[1] = 2.7;
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = 67;
            dr[1] = 2.3;
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = 65;
            dr[1] = 2.0;
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = 62;
            dr[1] = 1.7;
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = 60;
            dr[1] = 1.0;
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = 0;
            dr[1] = 0;
            dt.Rows.Add(dr);

            // add to dataset
            ds.Tables.Add(dt);
#endregion

#region 中科大算法
            dt = new DataTable() { TableName = "6" };

            // create the schema
            dt.Columns.Add("Score From");
            dt.Columns[0].DataType = Type.GetType("System.Single");
            dt.Columns.Add("GPA");
            dt.Columns[1].DataType = Type.GetType("System.Single");

            dr = dt.NewRow();
            dr[0] = 95;
            dr[1] = 4.3;
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = 90;
            dr[1] = 4.0;
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = 85;
            dr[1] = 3.7;
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = 82;
            dr[1] = 3.3;
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = 78;
            dr[1] = 3.0;
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = 75;
            dr[1] = 2.7;
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = 72;
            dr[1] = 2.3;
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = 68;
            dr[1] = 2.0;
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = 65;
            dr[1] = 1.7;
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = 64;
            dr[1] = 1.5;
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = 61;
            dr[1] = 1.3;
            dt.Rows.Add(dr);


            dr = dt.NewRow();
            dr[0] = 60;
            dr[1] = 1.0;
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = 0;
            dr[1] = 0;
            dt.Rows.Add(dr);

            // add to dataset
            ds.Tables.Add(dt);
#endregion
        }
    }
}
