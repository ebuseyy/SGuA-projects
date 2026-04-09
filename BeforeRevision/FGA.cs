using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;
using Microsoft.Office.Interop;
using Excel = Microsoft.Office.Interop.Excel;

namespace LCA
{
    public partial class FGA : Form
    {
        public FGA()
        {
            InitializeComponent();
        }



        double r;
        Random rastgele = new Random();
        double FunctionValue;
        int N = 18;

        double[] Alt = { -1, 0.2, 117.152, 1.6, 0.2, 523, 1.3, 0.2, 400, 1, 0.2, 500, 1, 0.2, 600, 1, 0.2, 400, 1 };
        double[] Ust = { -1, 1, 128.152, 2.2, 1, 575, 1.9, 1, 420, 1.6, 1, 510, 1.6, 1, 610, 1.6, 1, 420, 1.6 };
        double[,] populasyon = new double[7, 18 + 1];
        double[,] epopulasyon = new double[7, 18 + 1];
        double[,] kgpopulasyon = new double[7, 18 + 1];
        double[,] GlobalBest = new double[2, 18 + 2];
        double[,] siralipopulasyon = new double[19, 18 + 1];
        double[,] yenipopulasyon = new double[7, 18 + 1];

        private void button1_Click(object sender, EventArgs e)
        {
            
            //DataTable dtbpop = new DataTable();
            //dtbpop.Columns.Add("Fonksiyon Değeri", typeof(System.Double));
            //for (int i = 1; i <= 6; i++)
            //{
            //    dtbpop.Columns.Add("td" + i.ToString(), typeof(System.Double));
            //    dtbpop.Columns.Add("Ip" + i.ToString(), typeof(System.Double));
            //    dtbpop.Columns.Add("ti" + i.ToString(), typeof(System.Double));
            //}
            //DataTable dtepop = new DataTable();
            //dtepop.Columns.Add("Fonksiyon Değeri", typeof(System.Double));
            //for (int i = 1; i <= 6; i++)
            //{
            //    dtepop.Columns.Add("td" + i.ToString(), typeof(System.Double));
            //    dtepop.Columns.Add("Ip" + i.ToString(), typeof(System.Double));
            //    dtepop.Columns.Add("ti" + i.ToString(), typeof(System.Double));
            //}
            //DataTable dtkgpop = new DataTable();
            //dtkgpop.Columns.Add("Fonksiyon Değeri", typeof(System.Double));
            //for (int i = 1; i <= 6; i++)
            //{
            //    dtkgpop.Columns.Add("td" + i.ToString(), typeof(System.Double));
            //    dtkgpop.Columns.Add("Ip" + i.ToString(), typeof(System.Double));
            //    dtkgpop.Columns.Add("ti" + i.ToString(), typeof(System.Double));
            //}
            
            DataTable dttoplampop = new DataTable();
            dttoplampop.Columns.Add("Fonksiyon Değeri", typeof(System.Double));
            for (int i = 1; i <= 6; i++)
            {
                dttoplampop.Columns.Add("td" + i.ToString(), typeof(System.Double));
                dttoplampop.Columns.Add("Ip" + i.ToString(), typeof(System.Double));
                dttoplampop.Columns.Add("ti" + i.ToString(), typeof(System.Double));
            }
            Stopwatch sw = new Stopwatch();
            for (int calismasayisi = 0; calismasayisi < 10; calismasayisi++)
            {
                sw.Start();
                for (int i = 1; i < N + 2; i++)
                {
                    GlobalBest[1, i] = 1797693134862316;
                }
                baslangicpopulasyon();
                int Femax = Convert.ToInt32(textBox1.Text);
                for (int j = 0; j < Femax; j++)
                {
                    //dtbpop.Rows.Clear();
                    //dtepop.Rows.Clear();
                    //dtkgpop.Rows.Clear();
                    dttoplampop.Rows.Clear();
                    for (int ix = 1; ix <= 6; ix++)
                    {
                        //DataRow dr = dtbpop.NewRow();
                        //for (int jx = 0; jx <= 18; jx++)
                        //{
                        //    dr[jx] = populasyon[ix, jx].ToString();
                        //}
                        //dtbpop.Rows.Add(dr);
                        listBox1.Items.Add("Fonksiyon Değeri: "+populasyon[ix, 0].ToString() + "---td1: " + populasyon[ix, 1].ToString() + "---Ip1: " + populasyon[ix, 2].ToString() + "---ti1: " + populasyon[ix, 3].ToString() + "---td2: " + populasyon[ix, 4].ToString() + "---Ip2: " + populasyon[ix, 5].ToString() + "---ti2: " + populasyon[ix, 6].ToString() + "---td3: " + populasyon[ix, 7].ToString() + "---Ip3: " + populasyon[ix, 8].ToString() + "---ti3: " + populasyon[ix, 9].ToString() + "---td4: " + populasyon[ix, 10].ToString() + "---Ip4: " + populasyon[ix, 11].ToString() + "---ti4: " + populasyon[ix, 12].ToString() + "---td5: " + populasyon[ix, 13].ToString() + "---Ip5: " + populasyon[ix, 14].ToString() + "---ti5: " + populasyon[ix, 15].ToString() + "---td6: " + populasyon[ix, 16].ToString() + "---Ip6: " + populasyon[ix, 17].ToString() + "---ti6: " + populasyon[ix, 18].ToString());
                        
                    }
                    listBox1.Items.Add("-------------------------------------------------");
                    //dataGridView1.DataSource = dtbpop;



                    eslesme();
                    for (int ix = 1; ix <= 6; ix++)
                    {
                        //DataRow dr = dtepop.NewRow();
                        //for (int jx = 0; jx <= 18; jx++)
                        //{
                        //    dr[jx] = epopulasyon[ix, jx].ToString();
                        //}
                        //dtepop.Rows.Add(dr);
                        listBox2.Items.Add("Fonksiyon Değeri: " + epopulasyon[ix, 0].ToString() + "---td1: " + epopulasyon[ix, 1].ToString() + "---Ip1: " + epopulasyon[ix, 2].ToString() + "---ti1: " + epopulasyon[ix, 3].ToString() + "---td2: " + epopulasyon[ix, 4].ToString() + "---Ip2: " + epopulasyon[ix, 5].ToString() + "---ti2: " + epopulasyon[ix, 6].ToString() + "---td3: " + epopulasyon[ix, 7].ToString() + "---Ip3: " + epopulasyon[ix, 8].ToString() + "---ti3: " + epopulasyon[ix, 9].ToString() + "---td4: " + epopulasyon[ix, 10].ToString() + "---Ip4: " + epopulasyon[ix, 11].ToString() + "---ti4: " + epopulasyon[ix, 12].ToString() + "---td5: " + epopulasyon[ix, 13].ToString() + "---Ip5: " + epopulasyon[ix, 14].ToString() + "---ti5: " + epopulasyon[ix, 15].ToString() + "---td6: " + epopulasyon[ix, 16].ToString() + "---Ip6: " + epopulasyon[ix, 17].ToString() + "---ti6: " + epopulasyon[ix, 18].ToString());
                    }
                    listBox2.Items.Add("-------------------------------------------------");
                    //dataGridView2.DataSource = dtepop;



                    asilama();
                    for (int ix = 1; ix <= 6; ix++)
                    {
                        //DataRow dr = dtkgpop.NewRow();
                        //for (int jx = 0; jx <= 18; jx++)
                        //{
                        //    dr[jx] = kgpopulasyon[ix, jx].ToString();
                        //}
                        //dtkgpop.Rows.Add(dr);

                        listBox3.Items.Add("Fonksiyon Değeri: " + kgpopulasyon[ix, 0].ToString() + "---td1: " + kgpopulasyon[ix, 1].ToString() + "---Ip1: " + kgpopulasyon[ix, 2].ToString() + "---ti1: " + kgpopulasyon[ix, 3].ToString() + "---td2: " + kgpopulasyon[ix, 4].ToString() + "---Ip2: " + kgpopulasyon[ix, 5].ToString() + "---ti2: " + kgpopulasyon[ix, 6].ToString() + "---td3: " + kgpopulasyon[ix, 7].ToString() + "---Ip3: " + kgpopulasyon[ix, 8].ToString() + "---ti3: " + kgpopulasyon[ix, 9].ToString() + "---td4: " + kgpopulasyon[ix, 10].ToString() + "---Ip4: " + kgpopulasyon[ix, 11].ToString() + "---ti4: " + kgpopulasyon[ix, 12].ToString() + "---td5: " + kgpopulasyon[ix, 13].ToString() + "---Ip5: " + kgpopulasyon[ix, 14].ToString() + "---ti5: " + kgpopulasyon[ix, 15].ToString() + "---td6: " + kgpopulasyon[ix, 16].ToString() + "---Ip6: " + kgpopulasyon[ix, 17].ToString() + "---ti6: " + kgpopulasyon[ix, 18].ToString());
                        
                    }
                    listBox3.Items.Add("-------------------------------------------------");
                    //dataGridView3.DataSource = dtkgpop;

                    ///matrisleri birleştirme
                    for (int i = 1; i < 7; i++)
                    {
                        for (int a = 0; a < 19; a++)
                        {
                            siralipopulasyon[i, a] = populasyon[i, a];
                            siralipopulasyon[i + 6, a] = epopulasyon[i, a];
                            siralipopulasyon[i + 12, a] = kgpopulasyon[i, a];
                        }
                    }

                    for (int ix = 1; ix <= 18; ix++)
                    {
                        DataRow dr = dttoplampop.NewRow();
                        for (int jx = 0; jx <= 18; jx++)
                        {
                            dr[jx] = siralipopulasyon[ix, jx].ToString();
                        }
                        dttoplampop.Rows.Add(dr);
                    }
                    dttoplampop.DefaultView.Sort = "Fonksiyon Değeri asc";
                    dataGridView4.DataSource = dttoplampop;

                    for (int i = 1; i < 7; i++)
                    {
                        for (int a = 0; a < 19; a++)
                        {
                            yenipopulasyon[i, a] = (double)dataGridView4.Rows[i - 1].Cells[a].Value;
                        }
                    }
                    populasyon = kısıtlar(yenipopulasyon);

                    FunctionValue = populasyon[1, 3] + populasyon[1, 6] + populasyon[1, 9] + populasyon[1, 12] + populasyon[1, 15] + populasyon[1, 18];
                    for (int iy = 1; iy < N + 1; iy++)
                    {
                        GlobalBest[1, iy] = populasyon[1, iy];
                    }
                    GlobalBest[1, N + 1] = FunctionValue;

                }
                listBox4.Items.Add("..............." + Convert.ToInt32(calismasayisi+1) + ". çalışma ...........");
                listBox4.Items.Add("td-r1: " + GlobalBest[1, 1].ToString());
                listBox4.Items.Add("Ip-r1: " + GlobalBest[1, 2].ToString());
                listBox4.Items.Add("ti-r1: " + GlobalBest[1, 3].ToString());
                listBox4.Items.Add("td-r2: " + GlobalBest[1, 4].ToString());
                listBox4.Items.Add("Ip-r2: " + GlobalBest[1, 5].ToString());
                listBox4.Items.Add("ti-r2: " + GlobalBest[1, 6].ToString());
                listBox4.Items.Add("td-r3: " + GlobalBest[1, 7].ToString());
                listBox4.Items.Add("Ip-r3: " + GlobalBest[1, 8].ToString());
                listBox4.Items.Add("ti-r3: " + GlobalBest[1, 9].ToString());
                listBox4.Items.Add("td-r4: " + GlobalBest[1, 10].ToString());
                listBox4.Items.Add("Ip-r4: " + GlobalBest[1, 11].ToString());
                listBox4.Items.Add("ti-r4: " + GlobalBest[1, 12].ToString());
                listBox4.Items.Add("td-r5: " + GlobalBest[1, 13].ToString());
                listBox4.Items.Add("Ip-r5: " + GlobalBest[1, 14].ToString());
                listBox4.Items.Add("ti-r5: " + GlobalBest[1, 15].ToString());
                listBox4.Items.Add("td-r6: " + GlobalBest[1, 16].ToString());
                listBox4.Items.Add("Ip-r6: " + GlobalBest[1, 17].ToString());
                listBox4.Items.Add("ti-r6: " + GlobalBest[1, 18].ToString());

                listBox4.Items.Add("min ti: " + GlobalBest[1, N + 1].ToString());

                sw.Stop();
                listBox5.Items.Add(Convert.ToInt32(calismasayisi+1) + ". çalışma: "+sw.Elapsed.TotalSeconds.ToString());
                sw.Reset();
            }
        }
        public double[,] kısıtlar(double[,] NewTeamFormation)
        {
            ////////// KISITLAR /////////

            for (int i = 1; i < 7; i++)
            {
                double enbTD = 0;
                for (int z = 7; z < 19; z += 3)
                {
                    int roleno = (z + 2) / 3;
                    if (NewTeamFormation[i, z] < 0.2 || NewTeamFormation[i, z] > 1.0)
                    {
                        NewTeamFormation[i, z] = RandomNumberBetween(0.2, 1.0);
                        NewTeamFormation[i, z + 2] = Feval(NewTeamFormation[i, z], NewTeamFormation[i, z + 1], roleno);
                    }

                    if (roleno == 6)
                    {
                        if (NewTeamFormation[i, z + 1] < 400 || NewTeamFormation[i, z + 1] > 420)
                        {
                            NewTeamFormation[i, z + 1] = rastgele.NextDouble() * (420 - 400) + 400;
                            NewTeamFormation[i, z + 2] = Feval(NewTeamFormation[i, z], NewTeamFormation[i, z + 1], roleno);
                        }
                    }
                    else if (roleno == 5)
                    {
                        if (NewTeamFormation[i, z + 1] < 600 || NewTeamFormation[i, z + 1] > 610)
                        {
                            NewTeamFormation[i, z + 1] = rastgele.NextDouble() * (610 - 600) + 600;
                            NewTeamFormation[i, z + 2] = Feval(NewTeamFormation[i, z], NewTeamFormation[i, z + 1], roleno);
                        }
                    }
                    else if (roleno == 4)
                    {
                        if (NewTeamFormation[i, z + 1] < 500 || NewTeamFormation[i, z + 1] > 510)
                        {
                            NewTeamFormation[i, z + 1] = rastgele.NextDouble() * (510 - 500) + 500;
                            NewTeamFormation[i, z + 2] = Feval(NewTeamFormation[i, z], NewTeamFormation[i, z + 1], roleno);
                        }
                    }
                    else if (roleno == 3)
                    {
                        if (NewTeamFormation[i, z + 1] < 400 || NewTeamFormation[i, z + 1] > 420)
                        {
                            NewTeamFormation[i, z + 1] = rastgele.NextDouble() * (420 - 400) + 400;
                            NewTeamFormation[i, z + 2] = Feval(NewTeamFormation[i, z], NewTeamFormation[i, z + 1], roleno);
                        }
                    }

                    //NewTeamFormation[1, z + 1] = IpRasgele(roleno, NewTeamFormation[1, z + 1]);

                    if (NewTeamFormation[i, z + 2] < 1.0 || NewTeamFormation[i, z + 2] > 1.6)
                    {
                        NewTeamFormation[i, z + 2] = RandomNumberBetween(1.0, 1.6);
                        NewTeamFormation[i, z] = tdFeval(NewTeamFormation[i, z + 2], NewTeamFormation[i, z + 1], roleno);
                    }
                }
                for (int z = 7; z < N + 1; z += 3)
                {
                    if (enbTD < NewTeamFormation[i, z + 2])
                    {
                        enbTD = NewTeamFormation[i, z + 2];
                    }
                }

                if (NewTeamFormation[i, 4] < 0.2 || NewTeamFormation[i, 4] > 1.0)
                {
                    NewTeamFormation[i, 4] = RandomNumberBetween(0.2, 1.0);
                    NewTeamFormation[i, 6] = Feval(NewTeamFormation[i, 4], NewTeamFormation[i, 5], 2);
                }
                if (NewTeamFormation[i, 5] < 523 || NewTeamFormation[i, 5] > 575)
                {
                    NewTeamFormation[i, 5] = rastgele.NextDouble() * (575 - 523) + 523;
                    NewTeamFormation[i, 6] = Feval(NewTeamFormation[i, 4], NewTeamFormation[i, 5], 2);
                }
                //NewTeamFormation[1, 5] = IpRasgele(2, NewTeamFormation[1, 5]);
                if (enbTD + 0.3 > NewTeamFormation[i, 6] || NewTeamFormation[i, 6] > 1.9)
                {
                    NewTeamFormation[i, 6] = RandomNumberBetween(enbTD + 0.3, 1.9);
                    NewTeamFormation[i, 4] = tdFeval(NewTeamFormation[i, 6], NewTeamFormation[i, 5], 2);
                }

                //NewTeamFormation[1, 2] = IpRasgele(1,NewTeamFormation[1, 2]);

                if (NewTeamFormation[i, 1] < 0.2 || NewTeamFormation[i, 1] > 1.0)
                {
                    NewTeamFormation[i, 1] = RandomNumberBetween(0.2, 1.0);
                    NewTeamFormation[i, 3] = Feval(NewTeamFormation[i, 1], NewTeamFormation[i, 2], 1);
                }
                if (NewTeamFormation[i, 2] < 117.152 || NewTeamFormation[i, 2] > 128.65)
                {
                    NewTeamFormation[i, 2] = rastgele.NextDouble() * (128.65 - 117.52) + 117.52;
                    NewTeamFormation[i, 3] = Feval(NewTeamFormation[i, 1], NewTeamFormation[i, 2], 1);
                }

                if (NewTeamFormation[i, 6] + 0.3 > NewTeamFormation[i, 3] || NewTeamFormation[i, 3] > 2.2)
                {
                    NewTeamFormation[i, 3] = RandomNumberBetween(NewTeamFormation[i, 6] + 0.3, 2.2);
                    NewTeamFormation[i, 1] = tdFeval(NewTeamFormation[i, 3], NewTeamFormation[i, 2], 1);
                }


                NewTeamFormation[i, 0] = NewTeamFormation[i, 3] + NewTeamFormation[i, 6] + NewTeamFormation[i, 9] + NewTeamFormation[i, 12] + NewTeamFormation[i, 15] + NewTeamFormation[i, 18];
            }


            
            return NewTeamFormation;
        }
        public void baslangicpopulasyon()
        {
            ////////%%%%%%%%%%%%%%% Initialization %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
            /// Fidan 1

            populasyon[1, 1] = Alt[1];//td
            populasyon[1, 2] = Alt[2];//Ip
            populasyon[1, 3] = Feval(populasyon[1, 1], populasyon[1, 2], 1);//ti
            populasyon[1, 4] = Alt[4];
            populasyon[1, 5] = Alt[5];
            populasyon[1, 6] = Feval(populasyon[1, 4], populasyon[1, 5], 2);//ti
            populasyon[1, 7] = Alt[7];
            populasyon[1, 8] = Alt[8];
            populasyon[1, 9] = Feval(populasyon[1, 7], populasyon[1, 8], 3);//ti
            populasyon[1, 10] = Alt[10];
            populasyon[1, 11] = Alt[11];
            populasyon[1, 12] = Feval(populasyon[1, 10], populasyon[1, 11], 4);//ti
            populasyon[1, 13] = Alt[13];
            populasyon[1, 14] = Alt[14];
            populasyon[1, 15] = Feval(populasyon[1, 13], populasyon[1, 14], 5);//ti
            populasyon[1, 16] = Alt[16];
            populasyon[1, 17] = Alt[17];
            populasyon[1, 18] = Feval(populasyon[1, 16], populasyon[1, 17], 6);//ti

            FunctionValue = populasyon[1, 3] + populasyon[1, 6] + populasyon[1, 9] + populasyon[1, 12] + populasyon[1, 15] + populasyon[1, 18];
            if (FunctionValue < GlobalBest[1, N + 1])
            {
                for (int iy = 1; iy < N + 1; iy++)
                {
                    GlobalBest[1, iy] = populasyon[1, iy];
                }
                GlobalBest[1, N + 1] = FunctionValue;
            }
            populasyon[1, 0] = FunctionValue;

            /// Fidan 2
            populasyon[2, 1] = Ust[1];//td
            populasyon[2, 2] = Ust[2];//Ip
            populasyon[2, 3] = Feval(populasyon[2, 1], populasyon[2, 2], 1);//ti
            populasyon[2, 4] = Ust[4];
            populasyon[2, 5] = Ust[5];
            populasyon[2, 6] = Feval(populasyon[2, 4], populasyon[2, 5], 2);//ti
            populasyon[2, 7] = Ust[7];
            populasyon[2, 8] = Ust[8];
            populasyon[2, 9] = Feval(populasyon[2, 7], populasyon[2, 8], 3);//ti
            populasyon[2, 10] = Ust[10];
            populasyon[2, 11] = Ust[11];
            populasyon[2, 12] = Feval(populasyon[2, 10], populasyon[2, 11], 4);//ti
            populasyon[2, 13] = Ust[13];
            populasyon[2, 14] = Ust[14];
            populasyon[2, 15] = Feval(populasyon[2, 13], populasyon[2, 14], 5);//ti
            populasyon[2, 16] = Ust[16];
            populasyon[2, 17] = Ust[17];
            populasyon[2, 18] = Feval(populasyon[2, 16], populasyon[2, 17], 6);//ti

            FunctionValue = populasyon[2, 3] + populasyon[2, 6] + populasyon[2, 9] + populasyon[2, 12] + populasyon[2, 15] + populasyon[2, 18];
            if (FunctionValue < GlobalBest[1, N + 1])
            {
                for (int iy = 1; iy < N + 1; iy++)
                {
                    GlobalBest[1, iy] = populasyon[2, iy];
                }
                GlobalBest[1, N + 1] = FunctionValue;
            }
            populasyon[2, 0] = FunctionValue;

            /// Fidan 3
            r = rastgele.Next(0, 1) + rastgele.NextDouble();
            populasyon[3, 1] = Alt[1] + (Ust[1] - Alt[1]) * r;
            r = rastgele.Next(0, 1) + rastgele.NextDouble();
            populasyon[3, 2] = Alt[2] + (Ust[2] - Alt[2]) * r;
            populasyon[3, 3] = Feval(populasyon[3, 1], populasyon[3, 2], 1);//ti

            r = rastgele.Next(0, 1) + rastgele.NextDouble();
            populasyon[3, 4] = Alt[4] + (Ust[4] - Alt[4]) * r;
            r = rastgele.Next(0, 1) + rastgele.NextDouble();
            populasyon[3, 5] = Alt[5] + (Ust[5] - Alt[5]) * r;
            populasyon[3, 6] = Feval(populasyon[3, 4], populasyon[3, 5], 1);//ti

            r = rastgele.Next(0, 1) + rastgele.NextDouble();
            populasyon[3, 7] = Alt[7] + (Ust[7] - Alt[7]) * r;
            r = rastgele.Next(0, 1) + rastgele.NextDouble();
            populasyon[3, 8] = Alt[8] + (Ust[8] - Alt[8]) * r;
            populasyon[3, 9] = Feval(populasyon[3, 7], populasyon[3, 8], 1);//ti

            r = rastgele.Next(0, 1) + rastgele.NextDouble();
            populasyon[3, 10] = Alt[10] + (Ust[10] - Alt[10]) * r;
            r = rastgele.Next(0, 1) + rastgele.NextDouble();
            populasyon[3, 11] = Alt[11] + (Ust[11] - Alt[11]) * r;
            populasyon[3, 12] = Feval(populasyon[3, 10], populasyon[3, 11], 1);//ti

            r = rastgele.Next(0, 1) + rastgele.NextDouble();
            populasyon[3, 13] = Alt[13] + (Ust[13] - Alt[13]) * r;
            r = rastgele.Next(0, 1) + rastgele.NextDouble();
            populasyon[3, 14] = Alt[14] + (Ust[14] - Alt[14]) * r;
            populasyon[3, 15] = Feval(populasyon[3, 13], populasyon[3, 14], 1);//ti

            r = rastgele.Next(0, 1) + rastgele.NextDouble();
            populasyon[3, 16] = Alt[16] + (Ust[16] - Alt[16]) * r;
            r = rastgele.Next(0, 1) + rastgele.NextDouble();
            populasyon[3, 17] = Alt[17] + (Ust[17] - Alt[17]) * r;
            populasyon[3, 18] = Feval(populasyon[3, 16], populasyon[3, 17], 1);//ti

            FunctionValue = populasyon[3, 3] + populasyon[3, 6] + populasyon[3, 9] + populasyon[3, 12] + populasyon[3, 15] + populasyon[3, 18];
            if (FunctionValue < GlobalBest[1, N + 1])
            {
                for (int iy = 1; iy < N + 1; iy++)
                {
                    GlobalBest[1, iy] = populasyon[3, iy];
                }
                GlobalBest[1, N + 1] = FunctionValue;
            }
            populasyon[3, 0] = FunctionValue;

            /// Fidan 4
            r = rastgele.Next(0, 1) + rastgele.NextDouble();
            populasyon[4, 1] = Alt[1] + (Ust[1] - Alt[1]) * r;
            r = rastgele.Next(0, 1) + rastgele.NextDouble();
            populasyon[4, 2] = Alt[2] + (Ust[2] - Alt[2]) * r;
            populasyon[4, 3] = Feval(populasyon[4, 1], populasyon[4, 2], 1);//ti

            r = rastgele.Next(0, 1) + rastgele.NextDouble();
            populasyon[4, 4] = Alt[4] + (Ust[4] - Alt[4]) * r;
            r = rastgele.Next(0, 1) + rastgele.NextDouble();
            populasyon[4, 5] = Alt[5] + (Ust[5] - Alt[5]) * r;
            populasyon[4, 6] = Feval(populasyon[4, 4], populasyon[4, 5], 1);//ti

            r = rastgele.Next(0, 1) + rastgele.NextDouble();
            populasyon[4, 7] = Alt[7] + (Ust[7] - Alt[7]) * r;
            r = rastgele.Next(0, 1) + rastgele.NextDouble();
            populasyon[4, 8] = Alt[8] + (Ust[8] - Alt[8]) * r;
            populasyon[4, 9] = Feval(populasyon[4, 7], populasyon[4, 8], 1);//ti

            r = rastgele.Next(0, 1) + rastgele.NextDouble();
            populasyon[4, 10] = Alt[10] + (Ust[10] - Alt[10]) * r;
            r = rastgele.Next(0, 1) + rastgele.NextDouble();
            populasyon[4, 11] = Alt[11] + (Ust[11] - Alt[11]) * r;
            populasyon[4, 12] = Feval(populasyon[4, 10], populasyon[4, 11], 1);//ti

            r = rastgele.Next(0, 1) + rastgele.NextDouble();
            populasyon[4, 13] = Alt[13] + (Ust[13] - Alt[13]) * r;
            r = rastgele.Next(0, 1) + rastgele.NextDouble();
            populasyon[4, 14] = Alt[14] + (Ust[14] - Alt[14]) * r;
            populasyon[4, 15] = Feval(populasyon[4, 13], populasyon[4, 14], 1);//ti

            r = rastgele.Next(0, 1) + rastgele.NextDouble();
            populasyon[4, 16] = Alt[16] + (Ust[16] - Alt[16]) * (1 - r);
            r = rastgele.Next(0, 1) + rastgele.NextDouble();
            populasyon[4, 17] = Alt[17] + (Ust[17] - Alt[17]) * (1 - r);
            populasyon[4, 18] = Feval(populasyon[4, 16], populasyon[4, 17], 1);//ti

            FunctionValue = populasyon[4, 3] + populasyon[4, 6] + populasyon[4, 9] + populasyon[4, 12] + populasyon[4, 15] + populasyon[4, 18];
            if (FunctionValue < GlobalBest[1, N + 1])
            {
                for (int iy = 1; iy < N + 1; iy++)
                {
                    GlobalBest[1, iy] = populasyon[4, iy];
                }
                GlobalBest[1, N + 1] = FunctionValue;
            }
            populasyon[4, 0] = FunctionValue;

            /// Fidan 5
            r = rastgele.Next(0, 1) + rastgele.NextDouble();
            populasyon[5, 1] = Alt[1] + (Ust[1] - Alt[1]) * r;
            r = rastgele.Next(0, 1) + rastgele.NextDouble();
            populasyon[5, 2] = Alt[2] + (Ust[2] - Alt[2]) * r;
            populasyon[5, 3] = Feval(populasyon[5, 1], populasyon[5, 2], 1);//ti

            r = rastgele.Next(0, 1) + rastgele.NextDouble();
            populasyon[5, 4] = Alt[4] + (Ust[4] - Alt[4]) * r;
            r = rastgele.Next(0, 1) + rastgele.NextDouble();
            populasyon[5, 5] = Alt[5] + (Ust[5] - Alt[5]) * r;
            populasyon[5, 6] = Feval(populasyon[5, 4], populasyon[5, 5], 1);//ti

            r = rastgele.Next(0, 1) + rastgele.NextDouble();
            populasyon[5, 7] = Alt[7] + (Ust[7] - Alt[7]) * r;
            r = rastgele.Next(0, 1) + rastgele.NextDouble();
            populasyon[5, 8] = Alt[8] + (Ust[8] - Alt[8]) * r;
            populasyon[5, 9] = Feval(populasyon[5, 7], populasyon[5, 8], 1);//ti

            r = rastgele.Next(0, 1) + rastgele.NextDouble();
            populasyon[5, 10] = Alt[10] + (Ust[10] - Alt[10]) * r;
            r = rastgele.Next(0, 1) + rastgele.NextDouble();
            populasyon[5, 11] = Alt[11] + (Ust[11] - Alt[11]) * r;
            populasyon[5, 12] = Feval(populasyon[5, 10], populasyon[5, 11], 1);//ti

            r = rastgele.Next(0, 1) + rastgele.NextDouble();
            populasyon[5, 13] = Alt[13] + (Ust[13] - Alt[13]) * (1 - r);
            r = rastgele.Next(0, 1) + rastgele.NextDouble();
            populasyon[5, 14] = Alt[14] + (Ust[14] - Alt[14]) * (1 - r);
            populasyon[5, 15] = Feval(populasyon[5, 13], populasyon[5, 14], 1);//ti

            r = rastgele.Next(0, 1) + rastgele.NextDouble();
            populasyon[5, 16] = Alt[16] + (Ust[16] - Alt[16]) * r;
            r = rastgele.Next(0, 1) + rastgele.NextDouble();
            populasyon[5, 17] = Alt[17] + (Ust[17] - Alt[17]) * r;
            populasyon[5, 18] = Feval(populasyon[5, 16], populasyon[5, 17], 1);//ti

            FunctionValue = populasyon[5, 3] + populasyon[5, 6] + populasyon[5, 9] + populasyon[5, 12] + populasyon[5, 15] + populasyon[5, 18];
            if (FunctionValue < GlobalBest[1, N + 1])
            {
                for (int iy = 1; iy < N + 1; iy++)
                {
                    GlobalBest[1, iy] = populasyon[5, iy];
                }
                GlobalBest[1, N + 1] = FunctionValue;
            }
            populasyon[5, 0] = FunctionValue;

            /// Fidan 6
            r = rastgele.Next(0, 1) + rastgele.NextDouble();
            populasyon[6, 1] = Alt[1] + (Ust[1] - Alt[1]) * r;
            r = rastgele.Next(0, 1) + rastgele.NextDouble();
            populasyon[6, 2] = Alt[2] + (Ust[2] - Alt[2]) * r;
            populasyon[6, 3] = Feval(populasyon[6, 1], populasyon[6, 2], 1);//ti

            r = rastgele.Next(0, 1) + rastgele.NextDouble();
            populasyon[6, 4] = Alt[4] + (Ust[4] - Alt[4]) * r;
            r = rastgele.Next(0, 1) + rastgele.NextDouble();
            populasyon[6, 5] = Alt[5] + (Ust[5] - Alt[5]) * r;
            populasyon[6, 6] = Feval(populasyon[6, 4], populasyon[6, 5], 1);//ti

            r = rastgele.Next(0, 1) + rastgele.NextDouble();
            populasyon[6, 7] = Alt[7] + (Ust[7] - Alt[7]) * r;
            r = rastgele.Next(0, 1) + rastgele.NextDouble();
            populasyon[6, 8] = Alt[8] + (Ust[8] - Alt[8]) * r;
            populasyon[6, 9] = Feval(populasyon[6, 7], populasyon[6, 8], 1);//ti

            r = rastgele.Next(0, 1) + rastgele.NextDouble();
            populasyon[6, 10] = Alt[10] + (Ust[10] - Alt[10]) * r;
            r = rastgele.Next(0, 1) + rastgele.NextDouble();
            populasyon[6, 11] = Alt[11] + (Ust[11] - Alt[11]) * r;
            populasyon[6, 12] = Feval(populasyon[6, 10], populasyon[6, 11], 1);//ti

            r = rastgele.Next(0, 1) + rastgele.NextDouble();
            populasyon[6, 13] = Alt[13] + (Ust[13] - Alt[13]) * (1 - r);
            r = rastgele.Next(0, 1) + rastgele.NextDouble();
            populasyon[6, 14] = Alt[14] + (Ust[14] - Alt[14]) * (1 - r);
            populasyon[6, 15] = Feval(populasyon[6, 13], populasyon[6, 14], 1);//ti

            r = rastgele.Next(0, 1) + rastgele.NextDouble();
            populasyon[6, 16] = Alt[16] + (Ust[16] - Alt[16]) * (1 - r);
            r = rastgele.Next(0, 1) + rastgele.NextDouble();
            populasyon[6, 17] = Alt[17] + (Ust[17] - Alt[17]) * (1 - r);
            populasyon[6, 18] = Feval(populasyon[6, 16], populasyon[6, 17], 1);//ti

            FunctionValue = populasyon[6, 3] + populasyon[6, 6] + populasyon[6, 9] + populasyon[6, 12] + populasyon[6, 15] + populasyon[6, 18];
            if (FunctionValue < GlobalBest[1, N + 1])
            {
                for (int iy = 1; iy < N + 1; iy++)
                {
                    GlobalBest[1, iy] = populasyon[6, iy];
                }
                GlobalBest[1, N + 1] = FunctionValue;
            }
            populasyon[6, 0] = FunctionValue;
            ////////%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        }
        public void eslesme()
        {
            ////////%%%%%%%%%%%%%%% Eşleştirme %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

            //1-3 arası eşleştirme eşleştirme noktası 2
            epopulasyon[1, 1] = populasyon[1, 1];
            epopulasyon[1, 2] = populasyon[1, 2];
            epopulasyon[1, 3] = populasyon[1, 3];
            epopulasyon[1, 4] = populasyon[1, 4];
            epopulasyon[1, 5] = populasyon[1, 5];
            epopulasyon[1, 6] = populasyon[1, 6];
            epopulasyon[1, 7] = populasyon[3, 7];
            epopulasyon[1, 8] = populasyon[3, 8];
            epopulasyon[1, 9] = populasyon[3, 9];
            epopulasyon[1, 10] = populasyon[3, 10];
            epopulasyon[1, 11] = populasyon[3, 11];
            epopulasyon[1, 12] = populasyon[3, 12];
            epopulasyon[1, 13] = populasyon[3, 13];
            epopulasyon[1, 14] = populasyon[3, 14];
            epopulasyon[1, 15] = populasyon[3, 15];
            epopulasyon[1, 16] = populasyon[3, 16];
            epopulasyon[1, 17] = populasyon[3, 17];
            epopulasyon[1, 18] = populasyon[3, 18];

            FunctionValue = epopulasyon[1, 3] + epopulasyon[1, 6] + epopulasyon[1, 9] + epopulasyon[1, 12] + epopulasyon[1, 15] + epopulasyon[1, 18];
            if (FunctionValue < GlobalBest[1, N + 1])
            {
                for (int iy = 1; iy < N + 1; iy++)
                {
                    GlobalBest[1, iy] = epopulasyon[1, iy];
                }
                GlobalBest[1, N + 1] = FunctionValue;
            }
            epopulasyon[1, 0] = FunctionValue;

            epopulasyon[2, 1] = populasyon[3, 1];
            epopulasyon[2, 2] = populasyon[3, 2];
            epopulasyon[2, 3] = populasyon[3, 3];
            epopulasyon[2, 4] = populasyon[3, 4];
            epopulasyon[2, 5] = populasyon[3, 5];
            epopulasyon[2, 6] = populasyon[3, 6];
            epopulasyon[2, 7] = populasyon[1, 7];
            epopulasyon[2, 8] = populasyon[1, 8];
            epopulasyon[2, 9] = populasyon[1, 9];
            epopulasyon[2, 10] = populasyon[1, 10];
            epopulasyon[2, 11] = populasyon[1, 11];
            epopulasyon[2, 12] = populasyon[1, 12];
            epopulasyon[2, 13] = populasyon[1, 13];
            epopulasyon[2, 14] = populasyon[1, 14];
            epopulasyon[2, 15] = populasyon[1, 15];
            epopulasyon[2, 16] = populasyon[1, 16];
            epopulasyon[2, 17] = populasyon[1, 17];
            epopulasyon[2, 18] = populasyon[1, 18];

            FunctionValue = epopulasyon[2, 3] + epopulasyon[2, 6] + epopulasyon[2, 9] + epopulasyon[2, 12] + epopulasyon[2, 15] + epopulasyon[2, 18];
            if (FunctionValue < GlobalBest[1, N + 1])
            {
                for (int iy = 1; iy < N + 1; iy++)
                {
                    GlobalBest[1, iy] = epopulasyon[2, iy];
                }
                GlobalBest[1, N + 1] = FunctionValue;
            }
            epopulasyon[2, 0] = FunctionValue;

            //2-4 arası eşleştirme eşleştirme noktası 3
            epopulasyon[3, 1] = populasyon[4, 1];
            epopulasyon[3, 2] = populasyon[4, 2];
            epopulasyon[3, 3] = populasyon[4, 3];
            epopulasyon[3, 4] = populasyon[4, 4];
            epopulasyon[3, 5] = populasyon[4, 5];
            epopulasyon[3, 6] = populasyon[4, 6];
            epopulasyon[3, 7] = populasyon[4, 7];
            epopulasyon[3, 8] = populasyon[4, 8];
            epopulasyon[3, 9] = populasyon[4, 9];
            epopulasyon[3, 10] = populasyon[2, 10];
            epopulasyon[3, 11] = populasyon[2, 11];
            epopulasyon[3, 12] = populasyon[2, 12];
            epopulasyon[3, 13] = populasyon[2, 13];
            epopulasyon[3, 14] = populasyon[2, 14];
            epopulasyon[3, 15] = populasyon[2, 15];
            epopulasyon[3, 16] = populasyon[2, 16];
            epopulasyon[3, 17] = populasyon[2, 17];
            epopulasyon[3, 18] = populasyon[2, 18];

            FunctionValue = epopulasyon[3, 3] + epopulasyon[3, 6] + epopulasyon[3, 9] + epopulasyon[3, 12] + epopulasyon[3, 15] + epopulasyon[3, 18];
            if (FunctionValue < GlobalBest[1, N + 1])
            {
                for (int iy = 1; iy < N + 1; iy++)
                {
                    GlobalBest[1, iy] = epopulasyon[3, iy];
                }
                GlobalBest[1, N + 1] = FunctionValue;
            }
            epopulasyon[3, 0] = FunctionValue;

            epopulasyon[4, 1] = populasyon[2, 1];
            epopulasyon[4, 2] = populasyon[2, 2];
            epopulasyon[4, 3] = populasyon[2, 3];
            epopulasyon[4, 4] = populasyon[2, 4];
            epopulasyon[4, 5] = populasyon[2, 5];
            epopulasyon[4, 6] = populasyon[2, 6];
            epopulasyon[4, 7] = populasyon[2, 7];
            epopulasyon[4, 8] = populasyon[2, 8];
            epopulasyon[4, 9] = populasyon[2, 9];
            epopulasyon[4, 10] = populasyon[4, 10];
            epopulasyon[4, 11] = populasyon[4, 11];
            epopulasyon[4, 12] = populasyon[4, 12];
            epopulasyon[4, 13] = populasyon[4, 13];
            epopulasyon[4, 14] = populasyon[4, 14];
            epopulasyon[4, 15] = populasyon[4, 15];
            epopulasyon[4, 16] = populasyon[4, 16];
            epopulasyon[4, 17] = populasyon[4, 17];
            epopulasyon[4, 18] = populasyon[4, 18];

            FunctionValue = epopulasyon[4, 3] + epopulasyon[4, 6] + epopulasyon[4, 9] + epopulasyon[4, 12] + epopulasyon[4, 15] + epopulasyon[4, 18];
            if (FunctionValue < GlobalBest[1, N + 1])
            {
                for (int iy = 1; iy < N + 1; iy++)
                {
                    GlobalBest[1, iy] = epopulasyon[4, iy];
                }
                GlobalBest[1, N + 1] = FunctionValue;
            }
            epopulasyon[4, 0] = FunctionValue;

            //5-6 arası eşleştirme eşleştirme noktası 4
            epopulasyon[5, 1] = populasyon[5, 1];
            epopulasyon[5, 2] = populasyon[5, 2];
            epopulasyon[5, 3] = populasyon[5, 3];
            epopulasyon[5, 4] = populasyon[5, 4];
            epopulasyon[5, 5] = populasyon[5, 5];
            epopulasyon[5, 6] = populasyon[5, 6];
            epopulasyon[5, 7] = populasyon[5, 7];
            epopulasyon[5, 8] = populasyon[5, 8];
            epopulasyon[5, 9] = populasyon[5, 9];
            epopulasyon[5, 10] = populasyon[5, 10];
            epopulasyon[5, 11] = populasyon[5, 11];
            epopulasyon[5, 12] = populasyon[5, 12];
            epopulasyon[5, 13] = populasyon[6, 13];
            epopulasyon[5, 14] = populasyon[6, 14];
            epopulasyon[5, 15] = populasyon[6, 15];
            epopulasyon[5, 16] = populasyon[6, 16];
            epopulasyon[5, 17] = populasyon[6, 17];
            epopulasyon[5, 18] = populasyon[6, 18];

            FunctionValue = epopulasyon[5, 3] + epopulasyon[5, 6] + epopulasyon[5, 9] + epopulasyon[5, 12] + epopulasyon[5, 15] + epopulasyon[5, 18];
            if (FunctionValue < GlobalBest[1, N + 1])
            {
                for (int iy = 1; iy < N + 1; iy++)
                {
                    GlobalBest[1, iy] = epopulasyon[5, iy];
                }
                GlobalBest[1, N + 1] = FunctionValue;
            }
            epopulasyon[5, 0] = FunctionValue;

            epopulasyon[6, 1] = populasyon[6, 1];
            epopulasyon[6, 2] = populasyon[6, 2];
            epopulasyon[6, 3] = populasyon[6, 3];
            epopulasyon[6, 4] = populasyon[6, 4];
            epopulasyon[6, 5] = populasyon[6, 5];
            epopulasyon[6, 6] = populasyon[6, 6];
            epopulasyon[6, 7] = populasyon[6, 7];
            epopulasyon[6, 8] = populasyon[6, 8];
            epopulasyon[6, 9] = populasyon[6, 9];
            epopulasyon[6, 10] = populasyon[6, 10];
            epopulasyon[6, 11] = populasyon[6, 11];
            epopulasyon[6, 12] = populasyon[6, 12];
            epopulasyon[6, 13] = populasyon[5, 13];
            epopulasyon[6, 14] = populasyon[5, 14];
            epopulasyon[6, 15] = populasyon[5, 15];
            epopulasyon[6, 16] = populasyon[5, 16];
            epopulasyon[6, 17] = populasyon[5, 17];
            epopulasyon[6, 18] = populasyon[5, 18];

            FunctionValue = epopulasyon[6, 3] + epopulasyon[6, 6] + epopulasyon[6, 9] + epopulasyon[6, 12] + epopulasyon[6, 15] + epopulasyon[6, 18];
            if (FunctionValue < GlobalBest[1, N + 1])
            {
                for (int iy = 1; iy < N + 1; iy++)
                {
                    GlobalBest[1, iy] = epopulasyon[6, iy];
                }
                GlobalBest[1, N + 1] = FunctionValue;
            }
            epopulasyon[6, 0] = FunctionValue;
            ////////%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        }
        public void asilama()
        {
            ////////%%%%%%%%%%%%%%% Aşılama %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


            double[] r1dizi = new double[19];
            r = 0;
            //r nin hesabı
            for (int i = 1; i < 19; i++)
            {
                r = r + (Ust[i] - Alt[i]);
            }
            //r1 dizisinin hesabı
            double r1 = 0;
            for (int i = 1; i < 6; i++)
            {
                r1 = 0;
                for (int j = 1; j < 19; j++)
                {
                    r1 = r1 + (double)Math.Abs((populasyon[i, j] - populasyon[i + 1, j]));
                }
                r1dizi[i] = r1;
            }

            //karıştır geliştir 1-2 fidanlar
            kgpopulasyon[1, 1] = ((r1dizi[1] / r) * populasyon[1, 1]) + ((1 - (r1dizi[1] / r)) * populasyon[2, 1]);
            kgpopulasyon[1, 2] = ((r1dizi[1] / r) * populasyon[1, 2]) + ((1 - (r1dizi[1] / r)) * populasyon[2, 2]);
            kgpopulasyon[1, 3] = ((r1dizi[1] / r) * populasyon[1, 3]) + ((1 - (r1dizi[1] / r)) * populasyon[2, 3]);
            kgpopulasyon[1, 4] = ((r1dizi[1] / r) * populasyon[1, 4]) + ((1 - (r1dizi[1] / r)) * populasyon[2, 4]);
            kgpopulasyon[1, 5] = ((r1dizi[1] / r) * populasyon[1, 5]) + ((1 - (r1dizi[1] / r)) * populasyon[2, 5]);
            kgpopulasyon[1, 6] = ((r1dizi[1] / r) * populasyon[1, 6]) + ((1 - (r1dizi[1] / r)) * populasyon[2, 6]);
            kgpopulasyon[1, 7] = ((r1dizi[1] / r) * populasyon[1, 7]) + ((1 - (r1dizi[1] / r)) * populasyon[2, 7]);
            kgpopulasyon[1, 8] = ((r1dizi[1] / r) * populasyon[1, 8]) + ((1 - (r1dizi[1] / r)) * populasyon[2, 8]);
            kgpopulasyon[1, 9] = ((r1dizi[1] / r) * populasyon[1, 9]) + ((1 - (r1dizi[1] / r)) * populasyon[2, 9]);
            kgpopulasyon[1, 10] = ((r1dizi[1] / r) * populasyon[1, 10]) + ((1 - (r1dizi[1] / r)) * populasyon[2, 10]);
            kgpopulasyon[1, 11] = ((r1dizi[1] / r) * populasyon[1, 11]) + ((1 - (r1dizi[1] / r)) * populasyon[2, 11]);
            kgpopulasyon[1, 12] = ((r1dizi[1] / r) * populasyon[1, 12]) + ((1 - (r1dizi[1] / r)) * populasyon[2, 12]);
            kgpopulasyon[1, 13] = ((r1dizi[1] / r) * populasyon[1, 13]) + ((1 - (r1dizi[1] / r)) * populasyon[2, 13]);
            kgpopulasyon[1, 14] = ((r1dizi[1] / r) * populasyon[1, 14]) + ((1 - (r1dizi[1] / r)) * populasyon[2, 14]);
            kgpopulasyon[1, 15] = ((r1dizi[1] / r) * populasyon[1, 15]) + ((1 - (r1dizi[1] / r)) * populasyon[2, 15]);
            kgpopulasyon[1, 16] = ((r1dizi[1] / r) * populasyon[1, 16]) + ((1 - (r1dizi[1] / r)) * populasyon[2, 16]);
            kgpopulasyon[1, 17] = ((r1dizi[1] / r) * populasyon[1, 17]) + ((1 - (r1dizi[1] / r)) * populasyon[2, 17]);
            kgpopulasyon[1, 18] = ((r1dizi[1] / r) * populasyon[1, 18]) + ((1 - (r1dizi[1] / r)) * populasyon[2, 18]);

            FunctionValue = kgpopulasyon[1, 3] + kgpopulasyon[1, 6] + kgpopulasyon[1, 9] + kgpopulasyon[1, 12] + kgpopulasyon[1, 15] + kgpopulasyon[1, 18];
            if (FunctionValue < GlobalBest[1, N + 1])
            {
                for (int iy = 1; iy < N + 1; iy++)
                {
                    GlobalBest[1, iy] = kgpopulasyon[1, iy];
                }
                GlobalBest[1, N + 1] = FunctionValue;
            }
            kgpopulasyon[1, 0] = FunctionValue;

            kgpopulasyon[2, 1] = ((1 - (r1dizi[1] / r)) * populasyon[1, 1]) + ((r1dizi[1] / r) * populasyon[2, 1]);
            kgpopulasyon[2, 2] = ((1 - (r1dizi[1] / r)) * populasyon[1, 2]) + ((r1dizi[1] / r) * populasyon[2, 2]);
            kgpopulasyon[2, 3] = ((1 - (r1dizi[1] / r)) * populasyon[1, 3]) + ((r1dizi[1] / r) * populasyon[2, 3]);
            kgpopulasyon[2, 4] = ((1 - (r1dizi[1] / r)) * populasyon[1, 4]) + ((r1dizi[1] / r) * populasyon[2, 4]);
            kgpopulasyon[2, 5] = ((1 - (r1dizi[1] / r)) * populasyon[1, 5]) + ((r1dizi[1] / r) * populasyon[2, 5]);
            kgpopulasyon[2, 6] = ((1 - (r1dizi[1] / r)) * populasyon[1, 6]) + ((r1dizi[1] / r) * populasyon[2, 6]);
            kgpopulasyon[2, 7] = ((1 - (r1dizi[1] / r)) * populasyon[1, 7]) + ((r1dizi[1] / r) * populasyon[2, 7]);
            kgpopulasyon[2, 8] = ((1 - (r1dizi[1] / r)) * populasyon[1, 8]) + ((r1dizi[1] / r) * populasyon[2, 8]);
            kgpopulasyon[2, 9] = ((1 - (r1dizi[1] / r)) * populasyon[1, 9]) + ((r1dizi[1] / r) * populasyon[2, 9]);
            kgpopulasyon[2, 10] = ((1 - (r1dizi[1] / r)) * populasyon[1, 10]) + ((r1dizi[1] / r) * populasyon[2, 10]);
            kgpopulasyon[2, 11] = ((1 - (r1dizi[1] / r)) * populasyon[1, 11]) + ((r1dizi[1] / r) * populasyon[2, 11]);
            kgpopulasyon[2, 12] = ((1 - (r1dizi[1] / r)) * populasyon[1, 12]) + ((r1dizi[1] / r) * populasyon[2, 12]);
            kgpopulasyon[2, 13] = ((1 - (r1dizi[1] / r)) * populasyon[1, 13]) + ((r1dizi[1] / r) * populasyon[2, 13]);
            kgpopulasyon[2, 14] = ((1 - (r1dizi[1] / r)) * populasyon[1, 14]) + ((r1dizi[1] / r) * populasyon[2, 14]);
            kgpopulasyon[2, 15] = ((1 - (r1dizi[1] / r)) * populasyon[1, 15]) + ((r1dizi[1] / r) * populasyon[2, 15]);
            kgpopulasyon[2, 16] = ((1 - (r1dizi[1] / r)) * populasyon[1, 16]) + ((r1dizi[1] / r) * populasyon[2, 16]);
            kgpopulasyon[2, 17] = ((1 - (r1dizi[1] / r)) * populasyon[1, 17]) + ((r1dizi[1] / r) * populasyon[2, 17]);
            kgpopulasyon[2, 18] = ((1 - (r1dizi[1] / r)) * populasyon[1, 18]) + ((r1dizi[1] / r) * populasyon[2, 18]);

            FunctionValue = kgpopulasyon[2, 3] + kgpopulasyon[2, 6] + kgpopulasyon[2, 9] + kgpopulasyon[2, 12] + kgpopulasyon[2, 15] + kgpopulasyon[2, 18];
            if (FunctionValue < GlobalBest[1, N + 1])
            {
                for (int iy = 1; iy < N + 1; iy++)
                {
                    GlobalBest[1, iy] = kgpopulasyon[2, iy];
                }
                GlobalBest[1, N + 1] = FunctionValue;
            }
            kgpopulasyon[2, 0] = FunctionValue;

            //karıştır geliştir 3-4 fidanlar
            kgpopulasyon[3, 1] = ((r1dizi[1] / r) * populasyon[3, 1]) + ((1 - (r1dizi[1] / r)) * populasyon[4, 1]);
            kgpopulasyon[3, 2] = ((r1dizi[1] / r) * populasyon[3, 2]) + ((1 - (r1dizi[1] / r)) * populasyon[4, 2]);
            kgpopulasyon[3, 3] = ((r1dizi[1] / r) * populasyon[3, 3]) + ((1 - (r1dizi[1] / r)) * populasyon[4, 3]);
            kgpopulasyon[3, 4] = ((r1dizi[1] / r) * populasyon[3, 4]) + ((1 - (r1dizi[1] / r)) * populasyon[4, 4]);
            kgpopulasyon[3, 5] = ((r1dizi[1] / r) * populasyon[3, 5]) + ((1 - (r1dizi[1] / r)) * populasyon[4, 5]);
            kgpopulasyon[3, 6] = ((r1dizi[1] / r) * populasyon[3, 6]) + ((1 - (r1dizi[1] / r)) * populasyon[4, 6]);
            kgpopulasyon[3, 7] = ((r1dizi[1] / r) * populasyon[3, 7]) + ((1 - (r1dizi[1] / r)) * populasyon[4, 7]);
            kgpopulasyon[3, 8] = ((r1dizi[1] / r) * populasyon[3, 8]) + ((1 - (r1dizi[1] / r)) * populasyon[4, 8]);
            kgpopulasyon[3, 9] = ((r1dizi[1] / r) * populasyon[3, 9]) + ((1 - (r1dizi[1] / r)) * populasyon[4, 9]);
            kgpopulasyon[3, 10] = ((r1dizi[1] / r) * populasyon[3, 10]) + ((1 - (r1dizi[1] / r)) * populasyon[4, 10]);
            kgpopulasyon[3, 11] = ((r1dizi[1] / r) * populasyon[3, 11]) + ((1 - (r1dizi[1] / r)) * populasyon[4, 11]);
            kgpopulasyon[3, 12] = ((r1dizi[1] / r) * populasyon[3, 12]) + ((1 - (r1dizi[1] / r)) * populasyon[4, 12]);
            kgpopulasyon[3, 13] = ((r1dizi[1] / r) * populasyon[3, 13]) + ((1 - (r1dizi[1] / r)) * populasyon[4, 13]);
            kgpopulasyon[3, 14] = ((r1dizi[1] / r) * populasyon[3, 14]) + ((1 - (r1dizi[1] / r)) * populasyon[4, 14]);
            kgpopulasyon[3, 15] = ((r1dizi[1] / r) * populasyon[3, 15]) + ((1 - (r1dizi[1] / r)) * populasyon[4, 15]);
            kgpopulasyon[3, 16] = ((r1dizi[1] / r) * populasyon[3, 16]) + ((1 - (r1dizi[1] / r)) * populasyon[4, 16]);
            kgpopulasyon[3, 17] = ((r1dizi[1] / r) * populasyon[3, 17]) + ((1 - (r1dizi[1] / r)) * populasyon[4, 17]);
            kgpopulasyon[3, 18] = ((r1dizi[1] / r) * populasyon[3, 18]) + ((1 - (r1dizi[1] / r)) * populasyon[4, 18]);

            FunctionValue = kgpopulasyon[3, 3] + kgpopulasyon[3, 6] + kgpopulasyon[3, 9] + kgpopulasyon[3, 12] + kgpopulasyon[3, 15] + kgpopulasyon[3, 18];
            if (FunctionValue < GlobalBest[1, N + 1])
            {
                for (int iy = 1; iy < N + 1; iy++)
                {
                    GlobalBest[1, iy] = kgpopulasyon[3, iy];
                }
                GlobalBest[1, N + 1] = FunctionValue;
            }
            kgpopulasyon[3, 0] = FunctionValue;

            kgpopulasyon[4, 1] = ((1 - (r1dizi[3] / r)) * populasyon[3, 1]) + ((r1dizi[3] / r) * populasyon[4, 1]);
            kgpopulasyon[4, 2] = ((1 - (r1dizi[3] / r)) * populasyon[3, 2]) + ((r1dizi[3] / r) * populasyon[4, 2]);
            kgpopulasyon[4, 3] = ((1 - (r1dizi[3] / r)) * populasyon[3, 3]) + ((r1dizi[3] / r) * populasyon[4, 3]);
            kgpopulasyon[4, 4] = ((1 - (r1dizi[3] / r)) * populasyon[3, 4]) + ((r1dizi[3] / r) * populasyon[4, 4]);
            kgpopulasyon[4, 5] = ((1 - (r1dizi[3] / r)) * populasyon[3, 5]) + ((r1dizi[3] / r) * populasyon[4, 5]);
            kgpopulasyon[4, 6] = ((1 - (r1dizi[3] / r)) * populasyon[3, 6]) + ((r1dizi[3] / r) * populasyon[4, 6]);
            kgpopulasyon[4, 7] = ((1 - (r1dizi[3] / r)) * populasyon[3, 7]) + ((r1dizi[3] / r) * populasyon[4, 7]);
            kgpopulasyon[4, 8] = ((1 - (r1dizi[3] / r)) * populasyon[3, 8]) + ((r1dizi[3] / r) * populasyon[4, 8]);
            kgpopulasyon[4, 9] = ((1 - (r1dizi[3] / r)) * populasyon[3, 9]) + ((r1dizi[3] / r) * populasyon[4, 9]);
            kgpopulasyon[4, 10] = ((1 - (r1dizi[3] / r)) * populasyon[3, 10]) + ((r1dizi[3] / r) * populasyon[4, 10]);
            kgpopulasyon[4, 11] = ((1 - (r1dizi[3] / r)) * populasyon[3, 11]) + ((r1dizi[3] / r) * populasyon[4, 11]);
            kgpopulasyon[4, 12] = ((1 - (r1dizi[3] / r)) * populasyon[3, 12]) + ((r1dizi[3] / r) * populasyon[4, 12]);
            kgpopulasyon[4, 13] = ((1 - (r1dizi[3] / r)) * populasyon[3, 13]) + ((r1dizi[3] / r) * populasyon[4, 13]);
            kgpopulasyon[4, 14] = ((1 - (r1dizi[3] / r)) * populasyon[3, 14]) + ((r1dizi[3] / r) * populasyon[4, 14]);
            kgpopulasyon[4, 15] = ((1 - (r1dizi[3] / r)) * populasyon[3, 15]) + ((r1dizi[3] / r) * populasyon[4, 15]);
            kgpopulasyon[4, 16] = ((1 - (r1dizi[3] / r)) * populasyon[3, 16]) + ((r1dizi[3] / r) * populasyon[4, 16]);
            kgpopulasyon[4, 17] = ((1 - (r1dizi[3] / r)) * populasyon[3, 17]) + ((r1dizi[3] / r) * populasyon[4, 17]);
            kgpopulasyon[4, 18] = ((1 - (r1dizi[3] / r)) * populasyon[3, 18]) + ((r1dizi[3] / r) * populasyon[4, 18]);

            FunctionValue = kgpopulasyon[4, 3] + kgpopulasyon[4, 6] + kgpopulasyon[4, 9] + kgpopulasyon[4, 12] + kgpopulasyon[4, 15] + kgpopulasyon[4, 18];
            if (FunctionValue < GlobalBest[1, N + 1])
            {
                for (int iy = 1; iy < N + 1; iy++)
                {
                    GlobalBest[1, iy] = kgpopulasyon[4, iy];
                }
                GlobalBest[1, N + 1] = FunctionValue;
            }
            kgpopulasyon[4, 0] = FunctionValue;

            //karıştır geliştir 5-6 fidanlar
            kgpopulasyon[5, 1] = ((r1dizi[5] / r) * populasyon[5, 1]) + ((1 - (r1dizi[5] / r)) * populasyon[6, 1]);
            kgpopulasyon[5, 2] = ((r1dizi[5] / r) * populasyon[5, 2]) + ((1 - (r1dizi[5] / r)) * populasyon[6, 2]);
            kgpopulasyon[5, 3] = ((r1dizi[5] / r) * populasyon[5, 3]) + ((1 - (r1dizi[5] / r)) * populasyon[6, 3]);
            kgpopulasyon[5, 4] = ((r1dizi[5] / r) * populasyon[5, 4]) + ((1 - (r1dizi[5] / r)) * populasyon[6, 4]);
            kgpopulasyon[5, 5] = ((r1dizi[5] / r) * populasyon[5, 5]) + ((1 - (r1dizi[5] / r)) * populasyon[6, 5]);
            kgpopulasyon[5, 6] = ((r1dizi[5] / r) * populasyon[5, 6]) + ((1 - (r1dizi[5] / r)) * populasyon[6, 6]);
            kgpopulasyon[5, 7] = ((r1dizi[5] / r) * populasyon[5, 7]) + ((1 - (r1dizi[5] / r)) * populasyon[6, 7]);
            kgpopulasyon[5, 8] = ((r1dizi[5] / r) * populasyon[5, 8]) + ((1 - (r1dizi[5] / r)) * populasyon[6, 8]);
            kgpopulasyon[5, 9] = ((r1dizi[5] / r) * populasyon[5, 9]) + ((1 - (r1dizi[5] / r)) * populasyon[6, 9]);
            kgpopulasyon[5, 10] = ((r1dizi[5] / r) * populasyon[5, 10]) + ((1 - (r1dizi[5] / r)) * populasyon[6, 10]);
            kgpopulasyon[5, 11] = ((r1dizi[5] / r) * populasyon[5, 11]) + ((1 - (r1dizi[5] / r)) * populasyon[6, 11]);
            kgpopulasyon[5, 12] = ((r1dizi[5] / r) * populasyon[5, 12]) + ((1 - (r1dizi[5] / r)) * populasyon[6, 12]);
            kgpopulasyon[5, 13] = ((r1dizi[5] / r) * populasyon[5, 13]) + ((1 - (r1dizi[5] / r)) * populasyon[6, 13]);
            kgpopulasyon[5, 14] = ((r1dizi[5] / r) * populasyon[5, 14]) + ((1 - (r1dizi[5] / r)) * populasyon[6, 14]);
            kgpopulasyon[5, 15] = ((r1dizi[5] / r) * populasyon[5, 15]) + ((1 - (r1dizi[5] / r)) * populasyon[6, 15]);
            kgpopulasyon[5, 16] = ((r1dizi[5] / r) * populasyon[5, 16]) + ((1 - (r1dizi[5] / r)) * populasyon[6, 16]);
            kgpopulasyon[5, 17] = ((r1dizi[5] / r) * populasyon[5, 17]) + ((1 - (r1dizi[5] / r)) * populasyon[6, 17]);
            kgpopulasyon[5, 18] = ((r1dizi[5] / r) * populasyon[5, 18]) + ((1 - (r1dizi[5] / r)) * populasyon[6, 18]);

            FunctionValue = kgpopulasyon[5, 3] + kgpopulasyon[5, 6] + kgpopulasyon[5, 9] + kgpopulasyon[5, 12] + kgpopulasyon[5, 15] + kgpopulasyon[5, 18];
            if (FunctionValue < GlobalBest[1, N + 1])
            {
                for (int iy = 1; iy < N + 1; iy++)
                {
                    GlobalBest[1, iy] = kgpopulasyon[5, iy];
                }
                GlobalBest[1, N + 1] = FunctionValue;
            }
            kgpopulasyon[5, 0] = FunctionValue;

            kgpopulasyon[6, 1] = ((1 - (r1dizi[5] / r)) * populasyon[5, 1]) + ((r1dizi[5] / r) * populasyon[6, 1]);
            kgpopulasyon[6, 2] = ((1 - (r1dizi[5] / r)) * populasyon[5, 2]) + ((r1dizi[5] / r) * populasyon[6, 2]);
            kgpopulasyon[6, 3] = ((1 - (r1dizi[5] / r)) * populasyon[5, 3]) + ((r1dizi[5] / r) * populasyon[6, 3]);
            kgpopulasyon[6, 4] = ((1 - (r1dizi[5] / r)) * populasyon[5, 4]) + ((r1dizi[5] / r) * populasyon[6, 4]);
            kgpopulasyon[6, 5] = ((1 - (r1dizi[5] / r)) * populasyon[5, 5]) + ((r1dizi[5] / r) * populasyon[6, 5]);
            kgpopulasyon[6, 6] = ((1 - (r1dizi[5] / r)) * populasyon[5, 6]) + ((r1dizi[5] / r) * populasyon[6, 6]);
            kgpopulasyon[6, 7] = ((1 - (r1dizi[5] / r)) * populasyon[5, 7]) + ((r1dizi[5] / r) * populasyon[6, 7]);
            kgpopulasyon[6, 8] = ((1 - (r1dizi[5] / r)) * populasyon[5, 8]) + ((r1dizi[5] / r) * populasyon[6, 8]);
            kgpopulasyon[6, 9] = ((1 - (r1dizi[5] / r)) * populasyon[5, 9]) + ((r1dizi[5] / r) * populasyon[6, 9]);
            kgpopulasyon[6, 10] = ((1 - (r1dizi[5] / r)) * populasyon[5, 10]) + ((r1dizi[5] / r) * populasyon[6, 10]);
            kgpopulasyon[6, 11] = ((1 - (r1dizi[5] / r)) * populasyon[5, 11]) + ((r1dizi[5] / r) * populasyon[6, 11]);
            kgpopulasyon[6, 12] = ((1 - (r1dizi[5] / r)) * populasyon[5, 12]) + ((r1dizi[5] / r) * populasyon[6, 12]);
            kgpopulasyon[6, 13] = ((1 - (r1dizi[5] / r)) * populasyon[5, 13]) + ((r1dizi[5] / r) * populasyon[6, 13]);
            kgpopulasyon[6, 14] = ((1 - (r1dizi[5] / r)) * populasyon[5, 14]) + ((r1dizi[5] / r) * populasyon[6, 14]);
            kgpopulasyon[6, 15] = ((1 - (r1dizi[5] / r)) * populasyon[5, 15]) + ((r1dizi[5] / r) * populasyon[6, 15]);
            kgpopulasyon[6, 16] = ((1 - (r1dizi[5] / r)) * populasyon[5, 16]) + ((r1dizi[5] / r) * populasyon[6, 16]);
            kgpopulasyon[6, 17] = ((1 - (r1dizi[5] / r)) * populasyon[5, 17]) + ((r1dizi[5] / r) * populasyon[6, 17]);
            kgpopulasyon[6, 18] = ((1 - (r1dizi[5] / r)) * populasyon[5, 18]) + ((r1dizi[5] / r) * populasyon[6, 18]);

            FunctionValue = kgpopulasyon[6, 3] + kgpopulasyon[6, 6] + kgpopulasyon[6, 9] + kgpopulasyon[6, 12] + kgpopulasyon[6, 15] + kgpopulasyon[6, 18];
            if (FunctionValue < GlobalBest[1, N + 1])
            {
                for (int iy = 1; iy < N + 1; iy++)
                {
                    GlobalBest[1, iy] = kgpopulasyon[6, iy];
                }
                GlobalBest[1, N + 1] = FunctionValue;
            }
            kgpopulasyon[6, 0] = FunctionValue;
        }

        public double RandomNumberBetween(double minValue, double maxValue)
        {
            Random random = new Random();
            double next = random.NextDouble();
            return minValue + (next * (maxValue - minValue));
        }
        public double IpRasgele(int Rnum, double Ipx)
        {
            double deger = Ipx;
            if (Rnum == 6)
            {
                if (Ipx < 400 || Ipx > 420)
                {
                    deger = rastgele.NextDouble() * (420 - 400) + 400;
                }
            }
            else if (Rnum == 5)
            {
                if (Ipx < 600 || Ipx > 610)
                {
                    deger = rastgele.NextDouble() * (610 - 600) + 600;
                }
            }
            else if (Rnum == 4)
            {
                if (Ipx < 500 || Ipx > 510)
                {
                    deger = rastgele.NextDouble() * (510 - 500) + 500;
                }
            }
            else if (Rnum == 3)
            {
                if (Ipx < 400 || Ipx > 420)
                {
                    deger = rastgele.NextDouble() * (420 - 400) + 400;
                }
            }
            else if (Rnum == 2)
            {
                if (Ipx < 523 || Ipx > 575)
                {
                    deger = rastgele.NextDouble() * (575 - 523) + 523;
                }
            }
            else if (Rnum == 1)
            {
                if (Ipx < 117.152 || Ipx > 128.65)
                {
                    deger = rastgele.NextDouble() * (128.65 - 117.52) + 117.52;
                }
            }
            return deger;
        }
        public double Feval(double td, double lp, int i)
        {
            double FunctionValue = 0;
            //%%%%%%%%%%%%%%%%%%%%%%%%%%Over Current Relay%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
            if (i == 1)
            {
                FunctionValue = (0.14 * td) / ((Math.Pow((1263 / lp), 0.02)) - 1);
            }
            else
                FunctionValue = (0.14 * td) / ((Math.Pow((5639 / lp), 0.02)) - 1);

            return FunctionValue;
        }
        public double tdFeval(double ti, double lp, int i)
        {
            double td = 0;
            //%%%%%%%%%%%%%%%%%%%%%%%%%%Over Current Relay%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
            if (i == 1)
            {
                td = (ti * ((Math.Pow((1263 / lp), 0.02)) - 1)) / 0.14;
            }
            else
                td = (ti * ((Math.Pow((5639 / lp), 0.02)) - 1)) / 0.14;
            return td;
        }

        private void FGA_Load(object sender, EventArgs e)
        {
            this.Location = new Point(0, 0);
            //this.Size = new Size(1493, 761);

        }

        private void button2_Click(object sender, EventArgs e)
        {
            ExportToExcel(listBox4);
        }
        public static void ExportToExcel(ListBox ls)
        {
            Excel.Application objApp;
            Excel._Workbook objBook;
            Excel.Workbooks objBooks;
            Excel.Sheets objSheets;
            Excel._Worksheet objSheet;
            Excel.Range range;

            try
            {
                // Instantiate Excel and start a new workbook.
                objApp = new Excel.Application();
                objBooks = objApp.Workbooks;
                objBook = objBooks.Add(System.Reflection.Missing.Value);
                objSheets = objBook.Worksheets;
                objSheet = (Excel._Worksheet)objSheets.get_Item(1);
                range = objSheet.get_Range("A1", System.Reflection.Missing.Value);

                //fill the sheet
                int listCount = ls.Items.Count;
                range = range.get_Resize(listCount, 1);
                //Create an array.
                string[,] saRet = new string[listCount, 1];

                //Fill the array.
                for (int iRow = 0; iRow < listCount; iRow++)
                {
                    saRet[iRow, 0] = ls.Items[iRow].ToString();
                }

                range.set_Value(System.Reflection.Missing.Value, saRet);
                //range.set_Item(i, 0, listBox1.Items[i].ToString());
                objApp.Visible = true;
                objApp.UserControl = true;
            }
            catch (Exception theException)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, theException.Source);

                MessageBox.Show(errorMessage, "Error");
            }


        }

        private void button3_Click(object sender, EventArgs e)
        {
            ExportToExcel(listBox1);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            ExportToExcel(listBox2);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            ExportToExcel(listBox3);
        }
    }
}
