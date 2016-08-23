using System;
using System.Diagnostics;
using ELCUT;
using System.IO;
using System.Collections.Generic;
using System.Xml;

namespace InterfaceELCUT
{
    public static class Engine
    {
        private static Contour cnt;
        private static Model model;
        private static Problem task;
        private static ELCUT.Application app;
        private static string ProblemPath;
        private static FileInfo fi = new FileInfo("log.txt");

        /// <summary>
        /// Работа с логом
        /// </summary>
        private static void WriteToLog(string s)
        {
            StreamWriter sw;
            sw = fi.AppendText();
            sw.WriteLine(
                DateTime.Now.Hour + ":"
                + DateTime.Now.Minute + ":"
                + DateTime.Now.Second + " "
                + s);
            sw.Close();
        }

        /// <summary>
        /// Перезапуск модели при ошибке
        /// </summary>
        private static void ResetModel(Exception e)
        {
            WriteToLog("Ошибка: " + e.Message);
            app = new Application();
            task = (Problem)app.Problems.Open(ProblemPath);
            app.ActiveProblem.LoadModel();
            app.ActiveProblem.Model.Shapes.Blocks.BuildMesh(false, true);

            if (app.Models.Count > 0)
                model = (Model)app.Models.Item(1);
            app.MainWindow.WindowState = QfWindowState.qfMinimized;
            app.Windows.Item(1).WindowState = QfWindowState.qfMaximized;
        }

        /// <summary>
        /// Загрузить задачу.
        /// </summary>
        public static bool Load(string problemPath)
        {
            WriteToLog("----------------------------------------------------------");
            WriteToLog("Загрузка модели: " + problemPath);

            try
            {
                ProblemPath = problemPath;
                app = new Application();
                task = (Problem)app.Problems.Open(ProblemPath);
                app.ActiveProblem.LoadModel();
                if (app.Models.Count > 0)
                    model = (Model)app.Models.Item(1);
                app.MainWindow.WindowState = QfWindowState.qfMinimized;
                app.Windows.Item(1).WindowState = QfWindowState.qfMaximized;

                task = app.ActiveProblem;
                task.LoadModel();
                model = task.Model;
            }
            catch (Exception e)
            {
                ResetModel(e);
                Load(problemPath);
            }
            return true;
        }

        /// <summary>
        /// Включить/выключить магнит. 
        /// </summary>
        public static void EnableMagn(bool enabled, double coercitive)
        {
            WriteToLog("EnableMagn: enabled=" + enabled + ", coercive=" + coercitive);
            try
            {

                int magnCount = 0;
                while (true)
                {
                    try
                    {
                        Label magn = task.Labels[QfShapes.qfBlock].Item("МАГНИТ_" + (magnCount + 1));
                        magnCount++;
                    }
                    catch
                    {
                        break;
                    }
                }

                for (int i = 0; i < magnCount; i++)
                {
                    Label labelMagn = task.Labels[QfShapes.qfBlock].Item("МАГНИТ_" + (i + 1));
                    LabelBlockMS content = (LabelBlockMS)labelMagn.Content;
                    content.Coercive = app.PointRA(enabled ? coercitive : 0, 0);
                    labelMagn.Content = content;
                }

            }
            catch (Exception e)
            {
                ResetModel(e);
                EnableMagn(enabled, coercitive);
            }
        }

        /// <summary>
        /// Получить матрицу индуктивностей L для 3х фаз
        /// Выключаем магнит, подаем в 1 фазу статора номинальные токи и получаем потокосцепление для данной фазы
        /// Индуктивность вычисляем по формуле Psi = L * I
        /// </summary>
        public static double[,] GetLMatrix(double tokNominal)
        {

            WriteToLog("GetLMatrix: tokNominal=" + tokNominal);
            var L = new double[3, 3];

            try
            {
                task = app.ActiveProblem;
                task.LoadModel();
                model = task.Model;

                EnableMagn(false, 0);
                L = new double[3, 3];
                SetTokToFaza(1, tokNominal, true); //включаем малые токи
                SetTokToFaza(2, 0, true);
                SetTokToFaza(3, 0, true);
                task.SolveProblem(false);
                task.AnalyzeResults();

                for (int i = 0; i < 3; i++)
                    L[0, i] = Math.Abs(GetPsi(i + 1) / tokNominal); //получаем индуктивность при номинальных токах L=psi/I

                SetTokToFaza(1, 0, true);
                SetTokToFaza(2, tokNominal, true);
                SetTokToFaza(3, 0, true);
                task.SolveProblem(false);
                task.AnalyzeResults();

                for (int i = 0; i < 3; i++)
                    L[1, i] = Math.Abs(GetPsi(i + 1) / tokNominal);

                SetTokToFaza(1, 0, true);
                SetTokToFaza(2, 0, true);
                SetTokToFaza(3, tokNominal, true);
                task.SolveProblem(false);
                task.AnalyzeResults();

                for (int i = 0; i < 3; i++)
                    L[2, i] = Math.Abs(GetPsi(i + 1) / tokNominal);

            }
            catch (Exception e)
            {
                ResetModel(e);
                L = GetLMatrix(tokNominal);
            }
            return L;
        }





        /// <summary>
        /// Получить матрицу индуктивностей L для 9 фаз
        /// </summary>
        public static double[,] GetLMatrix9(double tokNominal)
        {
            var L = new double[9, 9];

            try
            {
                task = app.ActiveProblem;
                task.LoadModel();
                model = task.Model;
                L = new double[9, 9];
                EnableMagn(false, 0);

                var If = new double[9];
                for (int j = 0; j < 9; j++)
                {
                    for (int k = 0; k < 9; k++)
                        if (k == j)
                            If[k] = tokNominal;
                        else
                            If[k] = 0;

                    for (int x = 0; x < 9; x++)
                    {
                        SetTokToFaza(x+1, If[x], true);
                    }
                    

                    task.SolveProblem(false);
                    task.AnalyzeResults();
                    for (int i = 0; i < 9; i++)
                        L[j, i] = Math.Abs(GetPsi(i+1) / tokNominal);
                }
            }
            catch (Exception e)
            {
                ResetModel(e);
                L = GetLMatrix9(tokNominal);
            }
            return L;
        }







        /// <summary>
        /// Установить ток фазы A
        /// </summary>
        public static void SetTokToFaza(int fazaNum, double tokFaza, bool isTotal)
        {
            WriteToLog("SetTok: tokFaza=" + tokFaza + ", isTotal=" + isTotal);
            try
            {
                string lbl = "ФАЗА_" + fazaNum + "-";
                Label lb = task.Labels[QfShapes.qfBlock].Item(lbl);
                var cont = lb.Content as LabelBlockMS;
                cont.LoadingEx = Math.Round(-tokFaza, 2);
                cont.TotalCurrent = isTotal;
                cont.Serial = false;
                lb.Content = cont;

                lbl = "ФАЗА_" + fazaNum + "+";
                lb = task.Labels[QfShapes.qfBlock].Item(lbl);
                cont = lb.Content as LabelBlockMS;
                cont.LoadingEx = Math.Round(tokFaza, 2);
                cont.TotalCurrent = isTotal;
                cont.Serial = false;
                lb.Content = cont;

            }
            catch (Exception e)
            {
                ResetModel(e);
                SetTokToFaza(fazaNum, tokFaza, isTotal);
            }

        }


        /// <summary>
        /// Получаем потокосцепление для 1 фазы
        /// Потокосцепление считаем при поданом токе на данную фазу
        /// </summary>
        public static double GetPsi(int fazaNum)
        {
            WriteToLog("GetPhi: fazaNum=" + fazaNum);
            double res = 0;

            try
            {
                string lbla = "ФАЗА_" + fazaNum + "-";
                string lblb = "ФАЗА_" + fazaNum + "+";


                var wnd = task.Result.Windows.Item(1) as FieldWindow;
                Contour cnt = wnd.Contour;
                cnt.Delete(true);
                double resa, resb;
                cnt.AddBlock(lbla, null);
                resa = task.Result.GetIntegral(QfIntegrals.qfInt_FluxLinkage, cnt).Abs;
                cnt.Delete(true);
                cnt.AddBlock(lblb, null);
                resb = task.Result.GetIntegral(QfIntegrals.qfInt_FluxLinkage, cnt).Abs;
                cnt.Delete(true);
                res = (resb - resa); //суммируем потокосцепление с + и - фазы (- т.к. фаза resa отрицательная)
            }
            catch (Exception e)
            {
                ResetModel(e);
                res = GetPsi(fazaNum);
            }

            return res;
        }

        /// <summary>
        /// Получаем момент M
        /// </summary>
        public static double GetMoment()
        {
            WriteToLog("GetMoment()");
            double ret = 0;

            try
            {
                task.SolveProblem(false);
                task.AnalyzeResults();
                FieldWindow wnd = task.Result.Windows.Item(1) as FieldWindow;
                cnt = wnd.Contour;
                cnt.Delete(true);
                cnt.AddEdge("КОНТУР", null);
                ret = task.Result.GetIntegral(QfIntegrals.qfInt_MaxwellTorque, cnt).Abs;

            }
            catch (Exception e)
            {
                ResetModel(e);
                ret = GetMoment();
            }

            return ret;
        }

        /// <summary>
        /// Поворот ротора на угол phi
        /// </summary>
        public static void RotateMagn(double rad)
        {
            WriteToLog("RotateMagn: rad=" + rad);
            try
            {

                double initAngle = (task.Model.Shapes.get_LabeledAs("УГОЛ", "0", "0").Item(1) as Vertex).Point.Phi;

                Edge contur = task.Model.Shapes.get_LabeledAs("0", "КОНТУР", "0").Item(1) as Edge;
                ShapeRange rotor = task.Model.Shapes.get_InCircle(app.PointRA(0, 0), contur.Radius);
                rotor.Select();
                rotor.Move(QfTransformType.qfRotation, app.PointRA(0, 0), rad - initAngle);

                int magnCount = 0;

                while (true)
                {
                    try
                    {
                        Label magn = task.Labels[QfShapes.qfBlock].Item("МАГНИТ_" + (magnCount + 1));
                        magnCount++;
                    }
                    catch
                    {
                        break;
                    }
                }

                for (int i = 0; i < magnCount; i++)
                {
                    double angle = Math.PI * i;
                    Label labelMagn = task.Labels[QfShapes.qfBlock].Item("МАГНИТ_" + (i + 1));
                    LabelBlockMS content = (LabelBlockMS)labelMagn.Content;
                    content.Coercive = app.PointRA(content.Coercive.R, angle);
                    content.Polar = true;
                    labelMagn.Content = content;

                }

                app.ActiveProblem.Model.Shapes.Blocks.BuildMesh(false, true);
            }
            catch (Exception e)
            {
                ResetModel(e);
                RotateMagn(rad);
            }
        }


        /// <summary>
        /// Решить задачу.
        /// </summary>
        public static void Solve()
        {
            WriteToLog("Solve()");
            try
            {
                task.SolveProblem(false);
                task.AnalyzeResults();
            }
            catch (Exception e)
            {
                ResetModel(e);
                Solve();
            }
        }
    }
}
