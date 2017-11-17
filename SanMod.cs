using System;
using System.Windows.Forms;
using Ingeo;
using System.Runtime.InteropServices;
using System.Linq;
using System.Collections.Generic;
using System.IO;

namespace IngeoModules
{
    [ComVisible(true)]
    public class SanMod
    {
        public IIngeoApplication SanApp;

        public void Init(IIngeoApplication app)
        {
            SanApp = app;
        }

        public void Done()
        {
            SanApp = null;
        }

        public void ShowObjectInfo()
        {
            if (SanApp == null)
            {
                MessageBox.Show("Модуль не инициализирован");
            }
            else if ((SanApp.Selection.Count == 0) | (SanApp.Selection.Count > 1))
            {
                MessageBox.Show("Сначала выделите 1 объект");
            }
            else
            {
                string LeghtObject, SumString = " ";
                int TopoCount;
                IIngeoMapObject Per2 = SanApp.ActiveDb.MapObjects.GetObject(SanApp.Selection.get_IDs(0));
                LeghtObject = Per2.Perimeter.ToString();
                TopoCount = Per2.TopoLinks.Count;
                foreach (String topo in Per2.TopoLinks)
                {
                    SumString = SumString + topo + " ";
                }
                MessageBox.Show(String.Format("Объект : \n ID:{0} \n Периметр равен: {1} \n Количество топологических связей : {2} \n Связан с объектами с ID : {3}", SanApp.Selection.get_IDs(0), LeghtObject, TopoCount, SumString));
            }

        }   ///Показывает информацию о выделенном объекте
        public void ShowOffObjectsTwo()
        {
            if (SanApp.Selection.Count < 2)
            {
                MessageBox.Show("Необходимо выделить исходный и конечный объекты");
            }
            else
            {

                string[] startpoints = new string[2];

                int x;

                for (x = 0; x < 2; x++)
                {
                    startpoints[x] = SanApp.Selection.get_IDs(x);
                }
                IIngeoSelection selection = SanApp.Selection;
                //selection.DeselectAll();// Сбрасываем выделение всех объектов

                string idstart = startpoints[0], Wayname = "0";

                double PathCost = 0;

                int i, NumberOfId = 1, y = 0, h = 0, NumberOfStartObject = 0, q = 0, OutIDNumber = 0, FoundedObject = 0, f = 0;

                string[] IDs = new string[50000];
                for (i = 0; i < IDs.Length; i = i + 1) { IDs[i] = "0"; }
                IDs[0] = idstart;

                double[] Path = new double[50000];
                for (i = 0; i < Path.Length; i = i + 1) { Path[i] = 0; }
                Path[0] = 0;

                string[] Way = new string[50000];
                for (i = 0; i < Way.Length; i = i + 1) { Way[i] = ""; }

                IIngeoMapObject Per3, Per4;

                for (i = 1; i < 50000; i = i + 1)
                {
                    Per3 = SanApp.ActiveDb.MapObjects.GetObject(idstart);

                    for (h = 0; h < IDs.Length; h++)
                    {
                        if (idstart == IDs[h])
                        {
                            NumberOfStartObject = h; break;
                        }
                    }

                    foreach (String topoLink in Per3.TopoLinks)
                    {
                        Per4 = SanApp.ActiveDb.MapObjects.GetObject(topoLink);

                        PathCost = Per4.Perimeter;

                        Wayname = " " + topoLink;


                        foreach (String sectopolink in Per4.TopoLinks)
                        {
                            if ((IDs.Contains(sectopolink)) & (sectopolink != IDs[NumberOfStartObject]))
                            {
                                for (f = 0; f < IDs.Length; f++)
                                {
                                    if (sectopolink == IDs[f])
                                    {
                                        FoundedObject = f; break;
                                    }
                                }

                                if (Path[FoundedObject] > (PathCost + Path[NumberOfStartObject]))
                                {
                                    Path[FoundedObject] = PathCost + Path[NumberOfStartObject];
                                    Way[FoundedObject] = Wayname + Way[NumberOfStartObject];
                                }
                            }

                            if (!IDs.Contains(sectopolink))
                            {
                                IDs[NumberOfId] = sectopolink;
                                Path[NumberOfId] = PathCost + Path[NumberOfStartObject];
                                Way[NumberOfId] = Wayname + Way[NumberOfStartObject];
                                NumberOfId++;
                            }
                        }
                        y++;
                    }

                    if (IDs[i] != "0")
                    {
                        idstart = IDs[i];
                    }

                    else
                    {
                        break;
                    }

                }

                for (i = 0; i < 50000; ++i)
                {
                    if (IDs[i] == "0")
                    {
                        q = i;
                        break;

                    }
                }

                for (i = 0; i < 50000; ++i)
                {
                    if (IDs[i] == startpoints[1])
                    {
                        OutIDNumber = i;
                        break;
                    }
                }
                string AllObjectsIDsString = Way[OutIDNumber];
                string ClientsAdres = "";
                List<string> OutClientsAdreeses = FindClientsBySelectedWells(AllObjectsIDsString);
                foreach (string ClientsAdress in OutClientsAdreeses)
                {
                    ClientsAdres = ClientsAdres + " " + ClientsAdress;
                }
                MessageBox.Show(String.Format("Длинна отключенного участка : {0}; \nАдреса отключенных объектов:\n {1}", Path[OutIDNumber], ClientsAdres));
            }
        }    /// Трассировка без ответвлений с выводом длины отключенного участка
        public void ShowOffObjectsOne()
        {
            IIngeoMapObject Per1, Per2, NewPipeObject, Per5, Per6, Per7;
            IIngeoStyle SPer1;
            int cycleOne, cycleTwo; /// для циклов
            if (SanApp.Selection.Count < 2)
            {
                MessageBox.Show("Необходимо обозначить начальную \nи конечную точки участка трассировки");
                return;
            }
            //MessageBox.Show(String.Format("Определение отключенных объектов может занять некоторое время\n\n нажмите ОК для начала работы"));           
            else
            {
                string CurrentObjectID = SanApp.Selection.get_IDs(0); // берем id первого выделенного колодца
                string FinishPointID = SanApp.Selection.get_IDs(1); // берем id второго выделенного колодца
                int NumberOfId = 1, NumberOfCurrentObject = 0, q = 0, NumbersOfBranches = 1, MetkaOfFirstCycle = 1;

                string[] IDs = new string[50000]; /// массив найденых колодцев,можно заменить на динамичный массив, а так 50000 объектов максимум
                for (cycleOne = 0; cycleOne < IDs.Length; cycleOne = cycleOne + 1) { IDs[cycleOne] = "0"; } /// заполнили нулями
                IDs[0] = CurrentObjectID; ///первый из найденых колодцев конечно же выделенный нами
                int[] IDsNumberBranch = new int[50000]; /// массив номеров ветвей
                for (cycleOne = 0; cycleOne < IDsNumberBranch.Length; cycleOne = cycleOne + 1) { IDsNumberBranch[cycleOne] = 0; }
                int[] BranchEnd = new int[50000]; ///массив индикатора конца ветви
                for (cycleOne = 0; cycleOne < BranchEnd.Length; cycleOne = cycleOne + 1) { BranchEnd[cycleOne] = 0; }

                for (cycleOne = 1; cycleOne < 50000; cycleOne = cycleOne + 1) /// самый главный цикл автоматической трассировки сети
                {
                    IIngeoMapObject Well = SanApp.ActiveDb.MapObjects.GetObject(CurrentObjectID); /// в первый обход цикла подается сюда выделенная точка, в следующие обходы - найденные точки попорядку из массива IDs
                    for (cycleTwo = 0; cycleTwo < IDs.Length; cycleTwo++) ///    NumberOfCurrentObject присваивается индекс элемента из массива IDs
                    {
                        if (CurrentObjectID == IDs[cycleTwo])
                        {
                            NumberOfCurrentObject = cycleTwo;
                            break;
                        }
                    }
                    if (CurrentObjectID != FinishPointID)
                    {
                        foreach (String NewPipe in Well.TopoLinks)   /// перечисляем трубы от колодца
                        {
                            NewPipeObject = SanApp.ActiveDb.MapObjects.GetObject(NewPipe);
                            foreach (String NewWell in NewPipeObject.TopoLinks) /// перечисляем найденные колодцы на концах найденных труб
                            {
                                #region 
                                /* if ((IDs.Contains(NewWell)) )
                                  {
                                      for (j = 1; j < NumberOfStartObject; j = i + 1)
                                      {
                                          if (NewWell == IDs[j])
                                          {
                                              if (IDsNumberBranch[j] != IDsNumberBranch[NumberOfStartObject])
                                              {
                                                   goto CycleError;
                                              }

                                          }
                                      }

                                    /* IIngeoMapObject Per12 = SanApp.ActiveDb.MapObjects.GetObject(IDs[NumberOfStartObject]);

                                     if (Per12.TopoLinks.Count > 1)
                                     {

                                     }
                                }*/
                                #endregion  /// пытался заставить видеть циклические участки и оповещать об этом, времени не хватило
                                if (IDs[cycleOne] == FinishPointID)
                                {
                                    break;
                                }

                                if (!IDs.Contains(NewWell))    /// если найденного колодеца нет в массиве IDs
                                {
                                    IDs[NumberOfId] = NewWell; /// то добавляем
                                    if (MetkaOfFirstCycle == 1) /// если метка что основной цикл трассировки проходит в первый раз подтвержрается
                                    {
                                        IDsNumberBranch[NumberOfId] = NumbersOfBranches; /// то раздаем номера найденым ветвям
                                    }
                                    else /// если не в первый раз то найденным объектам присуждается номер ветви такой же как и у первоначальной точки от который она найдены (idstart)
                                    {
                                        IDsNumberBranch[NumberOfId] = IDsNumberBranch[NumberOfCurrentObject];
                                    }
                                    NumberOfId++; /// +1 индексу IDs
                                }
                            }
                            NumbersOfBranches++; /// в первый обход цикла эта переменная раздает номера ветвей
                        }
                    }
                    MetkaOfFirstCycle++; ///метка переключается 
                    if ((IDs[cycleOne] != "0")) /// если произведен перебор от всех найденных объектов то при обращении к следующему и натыкании на 0 цикл трассировки брикается , если не все перебраны то берется следующий элемент IDs
                    {
                        CurrentObjectID = IDs[cycleOne];
                    }
                    else
                    {
                        break;
                    }
                }

                for (cycleOne = 0; cycleOne < 50000; ++cycleOne) /// фактически q длинна IDs
                {
                    if (IDs[cycleOne] == "0")
                    {
                        q = cycleOne;
                        break;
                    }
                }

                for (cycleOne = 0; cycleOne < q; cycleOne++) /// если у колодца/пользователя только одна тополоническая связь то он  конечный
                {
                    Per1 = SanApp.ActiveDb.MapObjects.GetObject(IDs[cycleOne]);
                    if (Per1.TopoLinks.Count == 1) { BranchEnd[cycleOne] = 1; }
                }

                int[] SemData = new int[q];
                for (cycleOne = 0; cycleOne < q; cycleOne = cycleOne + 1) { SemData[cycleOne] = 0; }
                for (cycleOne = 0; cycleOne < q; cycleOne++) /// считываются диаметры крайних труб  и заносятся в Semdata
                {
                    if (BranchEnd[cycleOne] == 1)
                    {
                        Per2 = SanApp.ActiveDb.MapObjects.GetObject(IDs[cycleOne]);
                        foreach (string topolink in Per2.TopoLinks)
                        {
                            Per5 = SanApp.ActiveDb.MapObjects.GetObject(topolink);
                            SemData[cycleOne] = Convert.ToInt32(Per5.SemData.GetDisplayText("Трубы", "Диаметр", 0));
                        }
                    }
                }
                int[] NumberBranchOFF = new int[10];
                for (cycleOne = 0; cycleOne < 10; cycleOne = cycleOne + 1) { NumberBranchOFF[cycleOne] = 9; }
                for (cycleTwo = 1; cycleTwo < 10; cycleTwo++) /// определяется направление отключения от точки, если в определенной ветви  
                //нет разных диаметров конечных труб   то она помечается как отключаемая ветвь если же в ветви 
                //разные диаметры в том числе и большие то значит это бОльшая неотключаемая часть
                {
                    for (cycleOne = 0; cycleOne < q; cycleOne++)
                    {
                        if (IDsNumberBranch[cycleOne] == cycleTwo)
                        {
                            if (BranchEnd[cycleOne] == 1)
                            {
                                if ((1 <= SemData[cycleOne]) & (SemData[cycleOne] <= 299))
                                {
                                    if (NumberBranchOFF[cycleTwo] != 0)
                                        NumberBranchOFF[cycleTwo] = 1;
                                }
                                if ((300 <= SemData[cycleOne]) & (SemData[cycleOne] <= 1000))
                                {
                                    NumberBranchOFF[cycleTwo] = 0;
                                }
                            }
                        }
                    }
                }


                ShowFinalData();
                void ShowFinalData()
                {
                    string OutSrt = ""; /// строка выводящая информационное окно в конце обработки
                    List<string> OutPipes = new List<string>();
                    OutPipes.Add("Отключены трубы c кодами: ");
                    List<string> OutWells = new List<string>();
                    OutWells.Add("Отключены колодцы c кодами: ");
                    List<string> OutUsers = new List<string>();
                    OutUsers.Add("Отключены пользователи с ID: ");

                    string ObjectStyle = "";
                    for (cycleTwo = 0; cycleTwo < 10; cycleTwo++) /// вытаскиваем информацию по отключенным объектам
                    {
                        if (NumberBranchOFF[cycleTwo] == 1)
                        {
                            for (cycleOne = 0; cycleOne < q; cycleOne++)
                            {
                                if (IDsNumberBranch[cycleOne] == cycleTwo)
                                {
                                    OutSrt = OutSrt + " " + IDs[cycleOne];
                                    Per6 = SanApp.ActiveDb.MapObjects.GetObject(IDs[cycleOne]);
                                    foreach (string topolink in Per6.TopoLinks)
                                    {
                                        OutSrt = OutSrt + " " + topolink;
                                        Per7 = SanApp.ActiveDb.MapObjects.GetObject(topolink);
                                        foreach (IIngeoShape ser in Per6.Shapes)
                                        {
                                            SPer1 = ser.Style;
                                            ObjectStyle = SPer1.Name;
                                        }
                                        if (ObjectStyle == "Колодец")
                                        {
                                            OutWells.Add((Per6.SemData.GetDisplayText("Колодцы", "Код", 0)));
                                        }
                                        if (ObjectStyle == "Потребитель")
                                        {
                                            OutUsers.Add(topolink);
                                        }
                                        if (!OutPipes.Contains(topolink))
                                        {
                                            OutPipes.Add((Per7.SemData.GetDisplayText("Трубы", "Код", 0)));
                                            break;
                                        }
                                    }
                                }
                            }
                        }
                    }

                    SelectObjects(OutSrt);

                    StreamWriter writer = new StreamWriter(@"c:\diplom\off_objects.txt");
                    foreach (string s in OutPipes)
                    {
                        writer.WriteLine(s);
                    }
                    writer.Close();
                    StreamWriter writer2 = new StreamWriter(@"c:\diplom\off_objects.txt", true);
                    foreach (string s in OutWells)
                    {
                        writer2.WriteLine(s);
                    }
                    foreach (string s in OutUsers)
                    {
                        writer2.WriteLine(s);
                    }
                    writer2.Close();
                    string Out1 = (string.Join("  ", OutPipes.ToArray())) + "  ";
                    string Out2 = (string.Join("  ", OutWells.ToArray())) + "  ";
                    string Out3 = (string.Join("  ", OutUsers.ToArray())) + "  ";
                    MessageBox.Show(String.Format("Выделен участок попадающий в зону отключения.\n\n Информация по отключенному участку :\n\n{0}\n\n{1}\n\n{2} \n\n Информация по отключенным участкам занесена  на диск Е в файл Объекты отключения.txt   ", Out1, Out2, Out3));
                }
                List<string> OutClientsAdresses = FindClientsBySelectedWells("", true);
                string ClientsAdres = "";
                foreach (string ClientsAdress in OutClientsAdresses)
                {
                    ClientsAdres = ClientsAdres + " " + ClientsAdress;
                }
                MessageBox.Show(String.Format("Адреса отключенных объектов:\n{0}", ClientsAdres));
            }
        } //Трассировка с ответвлениями

        public void SelectObjects(string AllObjectsIDs)
        {
            int BeginDeleter = 0;
            char delimiterChars = ' ';
            string[] IdSplit = AllObjectsIDs.Split(delimiterChars); ///порезали AllObjectsIDs
            foreach (string o in IdSplit) /// выделили объекты
            {
                if (BeginDeleter != 0)
                {
                    SanApp.Selection.Select(o, 0);
                }
                BeginDeleter++;
            }
        }

        public List<string> FindClientsBySelectedWells(string AllObjectsIDsString = "", bool selected = false)
        {
            IIngeoMapObject MapObject, TopolinkMapObject, ClientMapObject;

            string ObjectStyle = "";
            var SelectionIDs = new List<string>(); ;
            if (selected == false) { SelectObjects(AllObjectsIDsString); }
            List<string> OutPipes = new List<string>();
            List<string> OutWells = new List<string>();
            List<string> OutUsers = new List<string>();
            List<string> ClientsPipesList = new List<string>();
            List<string> OutClientsIDs = new List<string>();
            List<string> OutClientsAdresses = new List<string>();
            IIngeoMapObjects mapobjs = SanApp.ActiveDb.MapObjects;
            //IIngeoMapObject Houses = SanApp.ActiveDb.MapObjects;

            for (int i = 0; i < SanApp.Selection.Count; i++)
            {
                SelectionIDs.Add(SanApp.Selection.IDs[i]);
            }
            string[] searchlayers = new string[] { "00010000042F", "0032000923B1", "00150000F2A6", "000100000DCC" };
            //поиск колодцев
            foreach (string Object in SelectionIDs)
            {
                MapObject = SanApp.ActiveDb.MapObjects.GetObject(Object);
                foreach (string Topolink in MapObject.TopoLinks)
                {
                    //OutSrt = OutSrt + " " + topolink;
                    TopolinkMapObject = SanApp.ActiveDb.MapObjects.GetObject(Topolink);
                    foreach (IIngeoShape ser in MapObject.Shapes)
                    {
                        IIngeoStyle SPer1 = ser.Style;
                        ObjectStyle = SPer1.Name;
                        if (ObjectStyle == "Колодец")
                        {
                            //OutWells.Add((MapObject.SemData.GetDisplayText("Колодцы", "Код", 0)));
                            if ((TopolinkMapObject.SemData.GetDisplayText("Трубы", "Назначение", 0) == "ВЫПУСК"))
                            {
                                foreach (IIngeoShape TopolinkShape in TopolinkMapObject.Shapes)
                                {
                                    IIngeoStyle TopolinkShapeStyle = TopolinkShape.Style;
                                    ObjectStyle = TopolinkShapeStyle.Name;
                                    IIngeoContour contour = TopolinkShape.Contour.BuildBufferZone(0.5);
                                    TIngeoContourRelation relation = TIngeoContourRelation.incrIntersected;
                                    TIngeoContourRelation relationmask = TIngeoContourRelation.incrIntersected | TIngeoContourRelation.incrTouched;
                                    IIngeoMapObjectsQuery qryObjects = mapobjs.QueryByContour(searchlayers, contour, relationmask, relation);

                                    string OutClientID = qryObjects.ObjectID;
                                    if (!OutClientsIDs.Contains(OutClientID))
                                    {
                                        ClientMapObject = SanApp.ActiveDb.MapObjects.GetObject(OutClientID);
                                        OutClientsIDs.Add(OutClientID);
                                        try
                                        {
                                            //алгоритм для домов. если такой семантики нет, то выдаст ексепшн и попробует алгоритм для сооружений
                                            ClientMapObject.SemData.GetDisplayText("Адрес", "Улица", 0);
                                            OutClientsAdresses.Add("Жилой дом:");
                                            OutClientsAdresses.Add((ClientMapObject.SemData.GetDisplayText("Адрес", "Улица", 0)));
                                            OutClientsAdresses.Add((ClientMapObject.SemData.GetDisplayText("Адрес", "номер дома", 0)));
                                            if (!string.IsNullOrEmpty((ClientMapObject.SemData.GetDisplayText("Адрес", "корпус", 0))))
                                            {
                                                OutClientsAdresses.Add("/");
                                                OutClientsAdresses.Add((ClientMapObject.SemData.GetDisplayText("Адрес", "корпус", 0)));
                                            }
                                            OutClientsAdresses.Add(";\n");
                                        }
                                        catch
                                        {
                                            try
                                            {
                                                ClientMapObject.SemData.GetDisplayText("Сооружения", "ID", 0);
                                                OutClientsAdresses.Add("Сооружение:");
                                                OutClientsAdresses.Add((ClientMapObject.SemData.GetDisplayText("Сооружения", "ID", 0)));
                                                OutClientsAdresses.Add((ClientMapObject.SemData.GetDisplayText("Сооружения", "Наименование", 0)));
                                                OutClientsAdresses.Add(";\n");
                                            }
                                            catch
                                            {
                                                try
                                                {
                                                    ClientMapObject.SemData.GetDisplayText("Бойлерная", "улица", 0);
                                                    OutClientsAdresses.Add("Бойлерная:");
                                                    OutClientsAdresses.Add((ClientMapObject.SemData.GetDisplayText("Бойлерная", "улица", 0)));
                                                    OutClientsAdresses.Add((ClientMapObject.SemData.GetDisplayText("Бойлерная", "номер дома", 0)));
                                                    if (!string.IsNullOrEmpty((ClientMapObject.SemData.GetDisplayText("Бойлерная", "корпус", 0))))
                                                    {
                                                        OutClientsAdresses.Add("/");
                                                        OutClientsAdresses.Add((ClientMapObject.SemData.GetDisplayText("Бойлерная", "корпус", 0)));
                                                    }
                                                    OutClientsAdresses.Add(";\n");
                                                }
                                                catch
                                                {

                                                }

                                            }

                                        }
                                    }
                                }
                            }
                        }
                        if (ObjectStyle == "Потребитель")
                        {
                            OutUsers.Add(Topolink);
                        }
                        if (!OutPipes.Contains(Topolink))
                        {
                            //OutPipes.Add((TopolinkMapObject.SemData.GetDisplayText("Трубы", "Код", 0)));
                            break;
                        }
                    }
                }
            }
            return OutClientsAdresses;
        }

    }
}


