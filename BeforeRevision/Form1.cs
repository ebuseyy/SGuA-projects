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
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        double[,] GlobalBestofEachRun;
        double[,] NumberofFunctionsEvaluatedInEachRun;
        double[,] ExecutionTimeofEachRun;
        long[,] ScoreBoard;
        double FunctionValue;
        Random rastgele = new Random();

        private void button1_Click(object sender, EventArgs e)
        {

            Stopwatch sw = new Stopwatch();

            for (int NumberOfRuns = 1; NumberOfRuns <= 10; NumberOfRuns++)
            {
                sw.Start();
                //NumberOfRuns
                int N = 18;
                int[] LowerBound = new int[N + 1];
                int[] UpperBound = new int[N + 1];
                for (int i = 1; i < N + 1; i++)
                {
                    LowerBound[i] = N * 1;
                    UpperBound[i] = N * 1;
                }

                int LeagueSize = Convert.ToInt32(textBox1.Text);
                double FEmax = Convert.ToInt32(textBox2.Text);
                double c1 = Convert.ToDouble(textBox3.Text);
                double c2 = Convert.ToDouble(textBox4.Text);
                double Pc = Convert.ToDouble(textBox5.Text);

                int TypeOfFormationUpdate = Convert.ToInt32(comboBox1.Text);

                double NumberOfSeasons = Math.Ceiling((double)(FEmax / (LeagueSize * (LeagueSize - 1))));
                double MaxWeeks = Math.Ceiling(Convert.ToDouble((LeagueSize - 1) * NumberOfSeasons));
                int StoppingFlag = 0;
                int FE = 0;

                GlobalBestofEachRun = new double[LeagueSize / 2 + 1, 2];
                NumberofFunctionsEvaluatedInEachRun = new double[LeagueSize / 2 + 1, 2];
                ExecutionTimeofEachRun = new double[LeagueSize / 2 + 1, 2];
                //%%%%%%%%%%%%%%%%%%%%% Generating League Time Table %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
                int[,] Timetable = new int[LeagueSize + 1, LeagueSize + 2];
                for (int i = 1; i < LeagueSize + 1; i++)
                {
                    Timetable[i, 1] = i;
                }
                int[,] BisectionedList = new int[3, LeagueSize / 2 + 1];
                for (int i = 1; i < LeagueSize / 2; i++)
                {
                    BisectionedList[1, i] = i;
                    BisectionedList[2, i + 1] = LeagueSize - i * 1;
                    if (i == 1)
                    {
                        BisectionedList[2, i] = LeagueSize;
                        BisectionedList[1, LeagueSize / 2] = LeagueSize / 2;
                    }
                }
                int[,] TemporaryList = new int[3, LeagueSize / 2 + 1];
                for (int i = 1; i <= LeagueSize; i++)
                {
                    for (int j = 1; j <= LeagueSize / 2; j++)
                    {
                        Timetable[BisectionedList[1, j], i + 1] = BisectionedList[2, j];
                        Timetable[BisectionedList[2, j], i + 1] = BisectionedList[1, j];
                    }

                    TemporaryList[1, 1] = 1;
                    TemporaryList[1, 2] = BisectionedList[2, 1];
                    for (int k = 3; k <= LeagueSize / 2; k++)
                    {
                        TemporaryList[1, k] = BisectionedList[1, k - 1];
                    }
                    for (int k = 1; k <= LeagueSize / 2 - 1; k++)
                    {
                        TemporaryList[2, k] = BisectionedList[2, k + 1];
                    }
                    TemporaryList[2, LeagueSize / 2] = BisectionedList[1, LeagueSize / 2];
                    Array.Copy(TemporaryList, BisectionedList, LeagueSize + LeagueSize / 2 + 3);
                }
                int[,] LeagueSchedule = new int[LeagueSize + 1, LeagueSize];

                for (int ix = 1; ix < LeagueSize + 1; ix++)
                {
                    for (int iy = 1; iy < LeagueSize; iy++)
                    {
                        LeagueSchedule[ix, iy] = Timetable[ix, iy + 1];
                    }

                }
                ////////%%%%%%%%%%%%%%% datagrid %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

                if (LeagueSize > 2 && NumberOfRuns == 1)
                {
                    DataTable dt = new DataTable();
                    dt.Columns.Add("Takımlar/Haftalar", typeof(System.String));
                    for (int i = 1; i <= LeagueSize - 1; i++)
                    {
                        dt.Columns.Add("H" + i.ToString(), typeof(System.String));
                    }

                    for (int i = 1; i <= LeagueSize; i++)
                    {
                        DataRow dr = dt.NewRow();
                        for (int j = 1; j <= LeagueSize; j++)
                        {
                            dr[j - 1] = "T" + Timetable[i, j].ToString();
                        }
                        dt.Rows.Add(dr);
                    }

                    dataGridView1.DataSource = dt;
                }
                ////////%%%%%%%%%%%%%%% Initialization %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
                double[,] GlobalBest = new double[2, N + 2];
                for (int iy = 1; iy < N + 2; iy++)
                {
                    GlobalBest[1, iy] = 1797693134862316;
                }

                double[,] Population = new double[LeagueSize + 1, N + 1];
                for (int ix = 1; ix < LeagueSize + 1; ix++)
                {
                    for (int iy = N; iy >= 1; iy -= 1)
                    {
                        Population[ix, iy] = rastgele.Next(0, N) + rastgele.NextDouble();
                    }
                }

                double[,] FvalArray = new double[LeagueSize + 1, 2];
                for (int ix = 1; ix < LeagueSize + 1; ix++)
                {
                    FunctionValue = 0;
                    for (int iy = 1; iy < N; iy += 3)
                    {
                        Population[ix, iy + 2] = Feval(Population[ix, iy], Population[ix, iy + 1], ix);
                        FunctionValue += Population[ix, iy + 2];
                    }

                    FE = FE + 1;
                    FvalArray[ix, 1] = FunctionValue;
                    if (FunctionValue < GlobalBest[1, N + 1])
                    {
                        for (int iy = 1; iy < N + 1; iy++)
                        {
                            GlobalBest[1, iy] = Population[ix, iy];
                        }
                        GlobalBest[1, N + 1] = FunctionValue;
                    }
                }

                double[,] X = new double[LeagueSize + 1, N + 2];
                for (int ix = 1; ix < LeagueSize + 1; ix++)
                {
                    for (int iy = 1; iy < N + 1; iy++)
                    {
                        X[ix, iy] = Population[ix, iy];
                    }
                    X[ix, N + 1] = FvalArray[ix, 1];
                }

                double[,] B = new double[LeagueSize + 1, N + 2];
                for (int ix = 1; ix < LeagueSize + 1; ix++)
                {
                    for (int iy = 1; iy < N + 2; iy++)
                    {
                        B[ix, iy] = X[ix, iy];
                    }
                }
                ////////%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
                int t = 1;
                ScoreBoard = new long[LeagueSize + 1, 2];
                int ii = 0;
                int jj = 0;
                double Pit;
                while (FE <= FEmax)
                {
                    double LB = GlobalBest[1, N + 1];
                    for (int NumberOfMatch = 0; NumberOfMatch < LeagueSize / 2; NumberOfMatch++)
                    {
                        for (int ax = 1; ax < LeagueSize + 1; ax++)
                        {
                            if (ScoreBoard[ax, 1] == 0)
                            {
                                ii = ax;
                                break;
                            }
                        }
                        for (int ax = 1; ax < LeagueSize + 1; ax++)
                        {
                            if (LeagueSchedule[ax, t] == ii)
                            {
                                jj = ax;
                                break;
                            }
                        }
                        if (X[jj, X.GetLength(1) - 1] + X[ii, X.GetLength(1) - 1] - 2 * LB == 0)
                            Pit = 0.5;
                        else
                            Pit = (X[jj, X.GetLength(1) - 1] - LB) / (X[jj, X.GetLength(1) - 1] + X[ii, X.GetLength(1) - 1] - 2 * LB);

                        double rand = rastgele.NextDouble();
                        if (rand <= Pit)
                        {
                            ScoreBoard[ii, 1] = 3;
                            ScoreBoard[jj, 1] = -1;
                        }
                        else
                        {
                            ScoreBoard[jj, 1] = 3;
                            ScoreBoard[ii, 1] = -1;
                        }
                    }
                    if (t % (LeagueSize - 1) == 0)
                    {
                        int[,] temp = new int[LeagueSize + 1, LeagueSize];
                        temp = (int[,])LeagueSchedule.Clone();

                        int boyutY = LeagueSchedule.GetLength(1);
                        LeagueSchedule = new int[LeagueSize + 1, boyutY + LeagueSize - 1];

                        for (int ix = 1; ix < LeagueSize + 1; ix++)
                        {
                            for (int iy = 1; iy < LeagueSize; iy++)
                            {
                                LeagueSchedule[ix, iy] = temp[ix, iy];
                            }
                        }
                        for (int ix = 1; ix < LeagueSize + 1; ix++)
                        {
                            for (int iy = boyutY; iy < boyutY + LeagueSize - 1; iy++)
                            {
                                LeagueSchedule[ix, iy] = Timetable[ix, iy - boyutY + 2];
                            }
                        }

                    }

                    double[,] CurrentWeekFormations = new double[LeagueSize + 1, N + 2];
                    for (int TeamIndex = 1; TeamIndex < LeagueSize; TeamIndex++)
                    {

                        int i = TeamIndex;
                        int j = LeagueSchedule[i, t];
                        int L = LeagueSchedule[i, t + 1];
                        int k = LeagueSchedule[L, t];
                        double qi = Math.Ceiling(Math.Log(1 - (1 - Math.Pow((1 - Pc), N)) * rastgele.NextDouble()) / Math.Log(1 - Pc));
                        //qi = ceil(log(1 - (1 - (1 - Pc) ^ N) * rand) / log(1 - Pc));
                        int[,] PermutedIndices = new int[2, N + 1];

                        for (int iy = 1; iy < N + 1; iy++)
                        {
                            PermutedIndices[1, iy] = rastgele.Next(1, N);
                        }

                        double random = 1;

                        random *= rastgele.NextDouble();

                        ////////////////////////////////////////

                        int CD = PermutedIndices[1, (int)qi];
                        double[,] NewTeamFormation = new double[2, N + 1];
                        if (TypeOfFormationUpdate == 1)
                        {
                            if (ScoreBoard[i, 1] == 3 && ScoreBoard[L, 1] == 3) //%i was winner and L was winner
                            {
                                for (int iy = 1; iy < N + 1; iy++)
                                {
                                    NewTeamFormation[1, iy] = B[i, iy];
                                }
                                NewTeamFormation[1, CD] = B[i, CD] + (c1 * (X[i, CD] - X[k, CD]) * random + c1 * (X[i, CD] - X[j, CD]) * random);
                            }
                            else if (ScoreBoard[i, 1] == 3 && ScoreBoard[L, 1] == -1)// %i was winner and L was loser
                            {
                                for (int iy = 1; iy < N + 1; iy++)
                                {
                                    NewTeamFormation[1, iy] = B[i, iy];
                                }
                                NewTeamFormation[1, CD] = B[i, CD] + (c2 * (X[k, CD] - X[i, CD]) * random + c1 * (X[i, CD] - X[j, CD]) * random);
                            }
                            else if (ScoreBoard[i, 1] == -1 && ScoreBoard[L, 1] == 3)// %i was  loser and L was winner
                            {
                                for (int iy = 1; iy < N + 1; iy++)
                                {
                                    NewTeamFormation[1, iy] = B[i, iy];
                                }
                                NewTeamFormation[1, CD] = B[i, CD] + (c1 * (X[i, CD] - X[k, CD]) * random + c2 * (X[j, CD] - X[i, CD]) * random);
                            }
                            else if (ScoreBoard[i, 1] == -1 && ScoreBoard[L, 1] == -1)// %i was  loser and L was loser
                            {
                                for (int iy = 1; iy < N + 1; iy++)
                                {
                                    NewTeamFormation[1, iy] = B[i, iy];
                                }

                                NewTeamFormation[1, CD] = B[i, CD] + (c2 * (X[k, CD] - X[i, CD]) * random + c2 * (X[j, CD] - X[i, CD]) * random);

                            }
                        }
                        else if (TypeOfFormationUpdate == 2)
                        {
                            if (ScoreBoard[i, 1] == 3 && ScoreBoard[L, 1] == 3) //%i was winner and L was winner
                            {
                                for (int iy = 1; iy < N + 1; iy++)
                                {
                                    NewTeamFormation[1, iy] = B[i, iy];
                                }
                                NewTeamFormation[1, CD] = B[i, CD] + (c1 * (B[i, CD] - B[k, CD]) * random + c1 * (B[i, CD] - B[j, CD]) * random);
                            }
                            else if (ScoreBoard[i, 1] == 3 && ScoreBoard[L, 1] == -1)// %i was winner and L was loser
                            {
                                for (int iy = 1; iy < N + 1; iy++)
                                {
                                    NewTeamFormation[1, iy] = B[i, iy];
                                }
                                NewTeamFormation[1, CD] = B[i, CD] + (c2 * (B[k, CD] - B[i, CD]) * random + c1 * (B[i, CD] - B[j, CD]) * random);

                            }
                            else if (ScoreBoard[i, 1] == -1 && ScoreBoard[L, 1] == 3)// %i was  loser and L was winner
                            {
                                for (int iy = 1; iy < N + 1; iy++)
                                {
                                    NewTeamFormation[1, iy] = B[i, iy];
                                }
                                NewTeamFormation[1, CD] = B[i, CD] + (c1 * (B[i, CD] - B[k, CD]) * random + c2 * (B[j, CD] - B[i, CD]) * random);
                            }
                            else if (ScoreBoard[i, 1] == -1 && ScoreBoard[L, 1] == -1)// %i was  loser and L was loser
                            {
                                for (int iy = 1; iy < N + 1; iy++)
                                {
                                    NewTeamFormation[1, iy] = B[i, iy];
                                }

                                NewTeamFormation[1, CD] = B[i, CD] + (c2 * (B[k, CD] - B[i, CD]) * random + c2 * (B[j, CD] - B[i, CD]) * random);

                            }
                        }

                        ////////// KISITLAR /////////
                        double enbTD = 0;
                        for (int z = 7; z < N + 1; z += 3)
                        {
                            int roleno = (z + 2) / 3;
                            if (NewTeamFormation[1, z] < 0.2 || NewTeamFormation[1, z] > 1.0)
                            {
                                NewTeamFormation[1, z] = RandomNumberBetween(0.2, 1.0);
                                NewTeamFormation[1, z + 2] = Feval(NewTeamFormation[1, z], NewTeamFormation[1, z + 1], roleno);
                            }

                            if (roleno == 6)
                            {
                                if (NewTeamFormation[1, z + 1] < 400 || NewTeamFormation[1, z + 1] > 420)
                                {
                                    NewTeamFormation[1, z + 1] = rastgele.NextDouble() * (420 - 400) + 400;
                                    NewTeamFormation[1, z + 2] = Feval(NewTeamFormation[1, z], NewTeamFormation[1, z + 1], roleno);
                                }
                            }
                            else if (roleno == 5)
                            {
                                if (NewTeamFormation[1, z + 1] < 600 || NewTeamFormation[1, z + 1] > 610)
                                {
                                    NewTeamFormation[1, z + 1] = rastgele.NextDouble() * (610 - 600) + 600;
                                    NewTeamFormation[1, z + 2] = Feval(NewTeamFormation[1, z], NewTeamFormation[1, z + 1], roleno);
                                }
                            }
                            else if (roleno == 4)
                            {
                                if (NewTeamFormation[1, z + 1] < 500 || NewTeamFormation[1, z + 1] > 510)
                                {
                                    NewTeamFormation[1, z + 1] = rastgele.NextDouble() * (510 - 500) + 500;
                                    NewTeamFormation[1, z + 2] = Feval(NewTeamFormation[1, z], NewTeamFormation[1, z + 1], roleno);
                                }
                            }
                            else if (roleno == 3)
                            {
                                if (NewTeamFormation[1, z + 1] < 400 || NewTeamFormation[1, z + 1] > 420)
                                {
                                    NewTeamFormation[1, z + 1] = rastgele.NextDouble() * (420 - 400) + 400;
                                    NewTeamFormation[1, z + 2] = Feval(NewTeamFormation[1, z], NewTeamFormation[1, z + 1], roleno);
                                }
                            }

                            //NewTeamFormation[1, z + 1] = IpRasgele(roleno, NewTeamFormation[1, z + 1]);

                            if (NewTeamFormation[1, z + 2] < 1.0 || NewTeamFormation[1, z + 2] > 1.6)
                            {
                                NewTeamFormation[1, z + 2] = RandomNumberBetween(1.0, 1.6);
                                NewTeamFormation[1, z] = tdFeval(NewTeamFormation[1, z + 2], NewTeamFormation[1, z + 1], roleno);
                            }
                        }
                        for (int z = 7; z < N + 1; z += 3)
                        {
                            if (enbTD < NewTeamFormation[1, z + 2])
                            {
                                enbTD = NewTeamFormation[1, z + 2];
                            }
                        }

                        if (NewTeamFormation[1, 4] < 0.2 || NewTeamFormation[1, 4] > 1.0)
                        {
                            NewTeamFormation[1, 4] = RandomNumberBetween(0.2, 1.0);
                            NewTeamFormation[1, 6] = Feval(NewTeamFormation[1, 4], NewTeamFormation[1, 5], 2);
                        }
                        if (NewTeamFormation[1, 5] < 523 || NewTeamFormation[1, 5] > 575)
                        {
                            NewTeamFormation[1, 5] = rastgele.NextDouble() * (575 - 523) + 523;
                            NewTeamFormation[1, 6] = Feval(NewTeamFormation[1, 4], NewTeamFormation[1, 5], 2);
                        }
                        //NewTeamFormation[1, 5] = IpRasgele(2, NewTeamFormation[1, 5]);
                        if (enbTD + 0.3 > NewTeamFormation[1, 6] || NewTeamFormation[1, 6] > 1.9)
                        {
                            NewTeamFormation[1, 6] = RandomNumberBetween(enbTD + 0.3, 1.9);
                            NewTeamFormation[1, 4] = tdFeval(NewTeamFormation[1, 6], NewTeamFormation[1, 5], 2);
                        }

                        //NewTeamFormation[1, 2] = IpRasgele(1,NewTeamFormation[1, 2]);

                        if (NewTeamFormation[1, 1] < 0.2 || NewTeamFormation[1, 1] > 1.0)
                        {
                            NewTeamFormation[1, 1] = RandomNumberBetween(0.2, 1.0);
                            NewTeamFormation[1, 3] = Feval(NewTeamFormation[1, 1], NewTeamFormation[1, 2], 1);
                        }
                        if (NewTeamFormation[1, 2] < 117.152 || NewTeamFormation[1, 2] > 128.65)
                        {
                            NewTeamFormation[1, 2] = rastgele.NextDouble() * (128.65 - 117.52) + 117.52;
                            NewTeamFormation[1, 3] = Feval(NewTeamFormation[1, 1], NewTeamFormation[1, 2], 1);
                        }

                        if (NewTeamFormation[1, 6] + 0.3 > NewTeamFormation[1, 3] || NewTeamFormation[1, 3] > 2.2)
                        {
                            NewTeamFormation[1, 3] = RandomNumberBetween(NewTeamFormation[1, 6] + 0.3, 2.2);
                            NewTeamFormation[1, 1] = tdFeval(NewTeamFormation[1, 3], NewTeamFormation[1, 2], 1);
                        }


                        FunctionValue = 0;
                        for (int iy = 3; iy < N + 1; iy += 3)
                        {
                            FunctionValue += NewTeamFormation[1, iy];
                        }
                        FE = FE + 1;

                        if (FunctionValue < GlobalBest[1, N + 1])
                        {
                            for (int iy = 1; iy < N + 1; iy++)
                            {
                                GlobalBest[1, iy] = NewTeamFormation[1, iy];
                            }
                            GlobalBest[1, N + 1] = FunctionValue;
                            listBox1.Items.Add("BestObj = " + FunctionValue.ToString() + "   FE = " + FE.ToString());
                        }



                        //}
                        for (int iy = 1; iy < N + 1; iy++)
                        {
                            CurrentWeekFormations[TeamIndex, iy] = NewTeamFormation[1, iy];
                        }
                        CurrentWeekFormations[TeamIndex, CurrentWeekFormations.GetLength(1) - 1] = FunctionValue;

                        if (FunctionValue < B[TeamIndex, N + 1])
                        {
                            for (int iy = 1; iy < N + 1; iy++)
                            {
                                B[TeamIndex, iy] = NewTeamFormation[1, iy];
                            }
                            B[TeamIndex, N + 1] = FunctionValue;
                        }
                        if (FE >= FEmax)
                        {
                            StoppingFlag = 1;
                            break;
                        }
                    }
                    if (StoppingFlag == 1)
                        break;

                    ScoreBoard = new long[LeagueSize + 1, 2];

                    X = (double[,])CurrentWeekFormations.Clone();
                    t = t + 1;
                }

                GlobalBestofEachRun[GlobalBestofEachRun.GetLength(0) - 1, 1] = GlobalBest[1, N + 1];
                listBox2.Items.Add("..............." + NumberOfRuns.ToString() + ". çalışma ...........");
                listBox2.Items.Add("td-r1: " + GlobalBest[1, 1].ToString());
                listBox2.Items.Add("Ip-r1: " + GlobalBest[1, 2].ToString());
                listBox2.Items.Add("ti-r1: " + GlobalBest[1, 3].ToString());
                listBox2.Items.Add("td-r2: " + GlobalBest[1, 4].ToString());
                listBox2.Items.Add("Ip-r2: " + GlobalBest[1, 5].ToString());
                listBox2.Items.Add("ti-r2: " + GlobalBest[1, 6].ToString());
                listBox2.Items.Add("td-r3: " + GlobalBest[1, 7].ToString());
                listBox2.Items.Add("Ip-r3: " + GlobalBest[1, 8].ToString());
                listBox2.Items.Add("ti-r3: " + GlobalBest[1, 9].ToString());
                listBox2.Items.Add("td-r4: " + GlobalBest[1, 10].ToString());
                listBox2.Items.Add("Ip-r4: " + GlobalBest[1, 11].ToString());
                listBox2.Items.Add("ti-r4: " + GlobalBest[1, 12].ToString());
                listBox2.Items.Add("td-r5: " + GlobalBest[1, 13].ToString());
                listBox2.Items.Add("Ip-r5: " + GlobalBest[1, 14].ToString());
                listBox2.Items.Add("ti-r5: " + GlobalBest[1, 15].ToString());
                listBox2.Items.Add("td-r6: " + GlobalBest[1, 16].ToString());
                listBox2.Items.Add("Ip-r6: " + GlobalBest[1, 17].ToString());
                listBox2.Items.Add("ti-r6: " + GlobalBest[1, 18].ToString());

                listBox2.Items.Add("min ti: " + GlobalBest[1, N + 1].ToString());



                NumberofFunctionsEvaluatedInEachRun[NumberofFunctionsEvaluatedInEachRun.GetLength(0) - 1, 1] = FE;
                listBox3.Items.Add(FE);
                sw.Stop();
                ExecutionTimeofEachRun[ExecutionTimeofEachRun.GetLength(0) - 1, 1] = Convert.ToDouble(sw.ElapsedMilliseconds);
                listBox4.Items.Add(sw.Elapsed.TotalSeconds.ToString());
                sw.Reset();
            }

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

        private void Form1_Load(object sender, EventArgs e)
        {
            this.Location = new Point(20, 20);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ExportToExcel(listBox2);
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



    }
}
