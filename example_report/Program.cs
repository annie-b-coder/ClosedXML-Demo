using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace example_report
{
    /* структура для создания отчёта, которая содержит наименование производственной единицы и относящееся к ней оборудование*/
    public struct EquipmentDailyReport
    {
        public string UnitName { get; set; }
        public List<Report> Equipments { get; set; }
    }
    /* структура, содержащая информацию по конкретной единице оборудования () */
    public struct Report
    {
        /* наименование */
        public string Name { get; set; }
        /* список событий с подробностями за сутки */
        public List<Event> Events { get; set; }
        /* список событий за сутки */
        public List<DayEvent> DayEvents { get; set; }
        /* список событий за месяц */
        public List<MonthEvent> MonthEvents { get; set; }
    }

    /* структура, содержащая обобщенную информацию о событиях за сутки */
    public struct DayEvent
    {
        /* длительность */
        public int Duration { get; set; }
        /* количество */
        public int Count { get; set; }
        /* тип события */
        public string Type { get; set; }
    }
    /* структура, содержащая обобщенную информацию о событиях за месяц */
    public struct MonthEvent
    {
        /* длительность */
        public int Duration { get; set; }
        /* количество */
        public int Count { get; set; }
        /* тип события */
        public string Type { get; set; }
    }

    /* структура, содержащая подробную информацию о событиях за сутки */
    public struct Event
    {
        /* начало события */
        public DateTime SD { get; set; }
        /* конец события */
        public DateTime ED { get; set; }
        /* длительность */
        public int Duration { get; set; }
        /* тип события */
        public string Type { set; get; }
    }

    /* производственная единица */
    public class Unit
    {
        /* идентификатор в бд */
        public int Id { get; set; }
        /* наименование */
        public string Name { get; set; }
        /* оборудование, относящееся к производственной единице */
        public List<Equipment> Equipments { get; set; }
    }
    /* оборудование */
    public class Equipment
    {
        /* идентификатор в бд */
        public int Id { get; set; }
        /* наименование */
        public string Name { get; set; }
    }

    /* класс, описывающий информацию о событиях, происходящих с оборудованием */
    public class DataEquipment
    {
        /* идентификатор в бд */
        public int EqId { get; set; }
        /* начало события */
        public DateTime Start { get; set; }
        /* конец события */
        public DateTime End { get; set; }
        /* длительность события */
        public int Duration
        {
            get { return (int)(End - Start).TotalMinutes; }
        }
        /* тип события */
        public int PeriodType{ get; set; }       
    }

    class Program
    {      

        /// <summary>
        /// функция, результатом которой является файл xlsx, содержащий данные отчёта
        /// </summary>
        /// <param name="date">дата для формирования отчёта</param>
        public static void GetEquipmentDataDaily(DateTime date)
        {
            DateTime start = date;
            DateTime end = date.AddDays(1);
            DateTime SD = new DateTime(start.Year, start.Month, start.Day, 0, 0, 0);
            DateTime ED = new DateTime(end.Year, end.Month, end.Day, 0, 0, 0);

            #region формирование входных данных для демонстрации работы программы
            List<Unit> units = new List<Unit>() {

                new Unit()
                {
                    Id = 1, Name = "Unit1", Equipments = new List<Equipment>()
                    {
                      new Equipment() { Id = 1, Name = "Equipment1"},
                      new Equipment() { Id = 2, Name = "Equipment2"},
                    }
                },
                new Unit()
                {
                    Id = 2, Name = "Unit2", Equipments = new List<Equipment>()
                    {
                      new Equipment() { Id = 3, Name = "Equipment3"},
                      new Equipment() { Id = 4, Name = "Equipment4"},
                    }
                }
                };

            List<DataEquipment> data_ = new List<DataEquipment>()
            {
                new DataEquipment() { EqId = 1, Start = new DateTime(2020,01,01,5,30,00), End =  new DateTime(2020,01,01,6,40,00), PeriodType = 1 },
                new DataEquipment() { EqId = 1, Start = new DateTime(2020,01,01,2,30,00), End =  new DateTime(2020,01,01,3,10,00), PeriodType = 2 },
                new DataEquipment() { EqId = 1, Start = new DateTime(2020,01,01,2,38,00), End =  new DateTime(2020,01,01,3,15,00), PeriodType = 2 },

                new DataEquipment() { EqId = 2, Start = new DateTime(2020,01,01,5,30,00), End =  new DateTime(2020,01,01,6,40,00), PeriodType = 1 },
                new DataEquipment() { EqId = 2, Start = new DateTime(2020,01,01,2,30,00), End =  new DateTime(2020,01,01,3,10,00), PeriodType = 2 },

                new DataEquipment() { EqId = 3, Start = new DateTime(2020,01,01,5,15,00), End =  new DateTime(2020,01,01,6,40,00), PeriodType = 1 },
                new DataEquipment() { EqId = 3, Start = new DateTime(2020,01,01,4,30,00), End =  new DateTime(2020,01,01,7,10,00), PeriodType = 2 },

                new DataEquipment() { EqId = 4, Start = new DateTime(2020,01,01,5,15,00), End =  new DateTime(2020,01,01,6,40,00), PeriodType = 1 },
                new DataEquipment() { EqId = 4, Start = new DateTime(2020,01,01,4,30,00), End =  new DateTime(2020,01,01,7,10,00), PeriodType = 2 },
            };

            #endregion
            
            /* создание списка экземпляров ежедневного отчёта для всех ПЕ*/
            List<EquipmentDailyReport> report = new List<EquipmentDailyReport>();

            /* проход по циклу ПЕ */
            foreach (var unit in units)
            {
                /* создание экземпляра класса ежедневного отчёта для текущей ПЕ*/
                EquipmentDailyReport r = new EquipmentDailyReport();

                r.UnitName = unit.Name;

                List<Report> eq_r = new List<Report>();
                r.Equipments = eq_r;
                List<DataEquipment> data = new List<DataEquipment>();

                var equipments = unit.Equipments;

                /* проход по циклу оборудования на текущей производственной единице */
                foreach (var item in equipments)
                {
                    /* получение данных о событиях оборудования за выбранный период */
                    data = data_.Where(z => z.Start <= ED && z.End >= SD && z.EqId == item.Id).ToList();
                    if (data.Count != 0)
                    {
                        Report item_c = new Report();
                        item_c.Name = item.Name;
                        List<Event> events = new List<Event>();
                        List<DayEvent> dayEvents = new List<DayEvent>();
                        List<MonthEvent> monthEvents = new List<MonthEvent>();
                        foreach (var d in data)
                        {
                            Event ev = new Event();
                            ev.SD = d.Start;
                            ev.ED = d.End;
                            ev.Duration = d.Duration;
                            switch (d.PeriodType)
                            {
                                case 1:
                                    {
                                        ev.Type = "EventType1";
                                        break;
                                    }
                                case 2:
                                    {
                                        ev.Type = "EventType2";
                                        break;
                                    }
                            }
                            events.Add(ev);
                        }

                        if (data.Where(z => z.PeriodType == 2).ToList().Count != 0)
                        {
                            DayEvent de = new DayEvent();
                            de.Count = data.Where(z => z.PeriodType == 2).ToList().Count;
                            de.Type = "EventType2";
                            de.Duration = data.Where(z => z.PeriodType == 2).Sum(z => z.Duration);
                            dayEvents.Add(de);
                        }
                        if (data.Where(z => z.PeriodType == 1).ToList().Count != 0)
                        {
                            DayEvent de = new DayEvent();
                            de.Count = data.Where(z => z.PeriodType == 1).ToList().Count;
                            de.Type = "EventType1";
                            de.Duration = data.Where(z => z.PeriodType == 1).Sum(z => z.Duration);
                            dayEvents.Add(de);
                        }
                        data = data_.Where(z => z.Start <= ED && z.End >= new DateTime(SD.Year, SD.Month, 1, 0, 0, 0) && z.EqId == item.Id).ToList();

                        if (data.Where(z => z.PeriodType == 2).ToList().Count != 0)
                        {
                            MonthEvent me = new MonthEvent();
                            me.Count = data.Where(z => z.PeriodType == 2).ToList().Count;
                            me.Type = "EventType2";
                            me.Duration = data.Where(z => z.PeriodType == 2).Sum(z => z.Duration);
                            monthEvents.Add(me);
                        }
                        if (data.Where(z => z.PeriodType == 1).ToList().Count != 0)
                        {
                            MonthEvent me = new MonthEvent();
                            me.Count = data.Where(z => z.PeriodType == 1).ToList().Count;
                            me.Type = "EventType1";
                            me.Duration = data.Where(z => z.PeriodType == 1).Sum(z => z.Duration);
                            monthEvents.Add(me);
                        }
                        item_c.DayEvents = dayEvents;
                        item_c.MonthEvents = monthEvents;
                        item_c.Events = events;
                        eq_r.Add(item_c);
                    }
                }
                if (eq_r.Count != 0) report.Add(r);

            }

            /* формирование xlsx файла */
            XLWorkbook wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("table");
            ws.Column(1).Width = 18;
            ws.Column(2).Width = 18;
            ws.Column(3).Width = 14;
            ws.Column(4).Width = 26;

            ws.Cell(1, 1).Value = "Equipment Daily Report " + date.ToString("dd/MM/yyyy");

            var title_style = ws.Range(1, 1, 1, 4).Merge().Style;
            title_style.Font.Bold = true;
            title_style.Font.FontSize = 18;
            title_style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            title_style.Border.OutsideBorder = XLBorderStyleValues.Thin;

            int i = 2;
            /* проход по циклу отчётных данных по каждой ПЕ */
            foreach (var rep in report)
            {
                ws.Cell(i, 1).Value = rep.UnitName;
                var unit_style = ws.Range(i, 1, i, 4).Merge().Style;
                unit_style.Font.Bold = true;
                unit_style.Font.FontSize = 18;
                unit_style.Fill.BackgroundColor = XLColor.LightGray;
                unit_style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                unit_style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                int j = i + 1;
                /* проход по циклу отчётных данных по каждой единице оборудования */
                foreach (var item in rep.Equipments)
                {
                    ws.Cell(j, 1).Value = item.Name;
                    var item_style = ws.Range(j, 1, j, 4).Merge().Style;
                    item_style.Font.Bold = true;
                    item_style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    item_style.Fill.BackgroundColor = XLColor.PastelBlue;
                    item_style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                    int time = j + 1;
                    int tt = time;

                    ws.Cell(time, 1).Value = "Start";
                    ws.Cell(time, 1).Style.Font.Bold = true;
                    ws.Cell(time, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    ws.Cell(time, 1).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    ws.Cell(time, 2).Value = "End";
                    ws.Cell(time, 2).Style.Font.Bold = true;
                    ws.Cell(time, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    ws.Cell(time, 2).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    ws.Cell(time, 3).Value = "Duration";
                    ws.Cell(time, 3).Style.Font.Bold = true;
                    ws.Cell(time, 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    ws.Cell(time, 3).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    ws.Cell(time, 4).Value = "EventType";
                    ws.Cell(time, 4).Style.Font.Bold = true;
                    ws.Cell(time, 4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    ws.Cell(time, 4).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                    tt = time + 1;

                    foreach (var ev in item.Events)
                    {
                        ws.Cell(tt, 1).Value = ev.SD;
                        ws.Cell(tt, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        ws.Cell(tt, 1).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        ws.Cell(tt, 2).Value = ev.ED;
                        ws.Cell(tt, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        ws.Cell(tt, 2).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        TimeSpan times_ = TimeSpan.FromSeconds(ev.Duration);
                        ws.Cell(tt, 3).Value = times_.ToString(@"d\:hh\:mm\:ss");
                        ws.Cell(tt, 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        ws.Cell(tt, 3).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        ws.Cell(tt, 4).Value = ev.Type;
                        ws.Cell(tt, 4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        ws.Cell(tt, 4).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                        tt = tt + 1;
                    }
                    int itogo = tt;

                    var empty_style = ws.Range(tt, 1, tt, 4).Merge().Style;
                    empty_style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    ws.Cell(itogo, 1).Value = "Total Day";
                    var d_style = ws.Range(itogo, 1, itogo, 2).Merge().Style;
                    d_style.Font.Bold = true;
                    d_style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    d_style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    d_style.Font.FontColor = XLColor.DarkBlue;

                    ws.Cell(itogo, 3).Value = "Total Month";
                    var m_style = ws.Range(itogo, 3, itogo, 4).Merge().Style;
                    m_style.Font.Bold = true;
                    m_style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    m_style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    m_style.Font.FontColor = XLColor.DarkBlue;
                    ws.Cell(itogo, 3).Style.Font.Bold = true;
                    ws.Cell(itogo, 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    ws.Cell(itogo, 3).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    ws.Cell(itogo, 3).Style.Font.FontColor = XLColor.DarkBlue;
                    itogo++;

                    int itogo_ = itogo;

                    foreach (var ev in item.DayEvents)
                    {
                        ws.Cell(itogo, 1).Value = ev.Type;
                        ws.Range(itogo, 1, itogo, 4).Merge().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        ws.Range(itogo, 1, itogo, 4).Merge().Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        ws.Cell(itogo + 1, 1).Value = "Count";
                        ws.Cell(itogo + 1, 2).Value = ev.Count;
                        ws.Cell(itogo + 1, 1).Style.Font.Bold = true;
                        ws.Cell(itogo + 1, 2).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        ws.Cell(itogo + 1, 1).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        ws.Cell(itogo + 1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        ws.Cell(itogo + 1, 1).Style.Font.FontColor = XLColor.DarkBlue;
                        ws.Cell(itogo + 2, 1).Value = "Duration";
                        TimeSpan times_ = TimeSpan.FromSeconds(ev.Duration);
                        ws.Cell(itogo + 2, 2).Value = times_.ToString(@"d\:hh\:mm\:ss");
                        ws.Cell(itogo + 2, 1).Style.Font.Bold = true;
                        ws.Cell(itogo + 2, 2).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        ws.Cell(itogo + 2, 1).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        ws.Cell(itogo + 2, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        ws.Cell(itogo + 2, 1).Style.Font.FontColor = XLColor.DarkBlue;

                        itogo = itogo + 3;
                    }

                    if (item.DayEvents.Count < item.MonthEvents.Count)
                    {
                        ws.Range(itogo, 1, itogo, 4).Merge().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        ws.Range(itogo, 1, itogo, 4).Merge().Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        ws.Cell(itogo + 1, 1).Value = "Count";
                        ws.Cell(itogo + 1, 2).Value = 0;
                        ws.Cell(itogo + 1, 1).Style.Font.Bold = true;
                        ws.Cell(itogo + 1, 2).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        ws.Cell(itogo + 1, 1).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        ws.Cell(itogo + 1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        ws.Cell(itogo + 1, 1).Style.Font.FontColor = XLColor.DarkBlue;
                        ws.Cell(itogo + 2, 1).Value = "Duration";
                        ws.Cell(itogo + 2, 2).Value = 0;
                        ws.Cell(itogo + 2, 1).Style.Font.Bold = true;
                        ws.Cell(itogo + 2, 2).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        ws.Cell(itogo + 2, 1).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        ws.Cell(itogo + 2, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        ws.Cell(itogo + 2, 1).Style.Font.FontColor = XLColor.DarkBlue;
                        itogo = itogo + 3;
                    }
                    foreach (var ev in item.MonthEvents)
                    {
                        ws.Cell(itogo_, 1).Value = ev.Type;
                        ws.Range(itogo_, 1, itogo_, 4).Merge().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        ws.Range(itogo_, 1, itogo_, 4).Merge().Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        ws.Cell(itogo_ + 1, 3).Value = "Count";
                        ws.Cell(itogo_ + 1, 4).Value = ev.Count;
                        ws.Cell(itogo_ + 1, 3).Style.Font.Bold = true;
                        ws.Cell(itogo_ + 1, 4).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        ws.Cell(itogo_ + 1, 3).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        ws.Cell(itogo_ + 1, 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        ws.Cell(itogo_ + 1, 3).Style.Font.FontColor = XLColor.DarkBlue;
                        ws.Cell(itogo_ + 2, 3).Value = "Duration";
                        TimeSpan times_ = TimeSpan.FromSeconds(ev.Duration);
                        ws.Cell(itogo_ + 2, 4).Value = times_.ToString(@"d\:hh\:mm\:ss");
                        ws.Cell(itogo_ + 2, 3).Style.Font.Bold = true;
                        ws.Cell(itogo_ + 2, 4).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        ws.Cell(itogo_ + 2, 3).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        ws.Cell(itogo_ + 2, 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        ws.Cell(itogo_ + 2, 3).Style.Font.FontColor = XLColor.DarkBlue;
                        itogo_ = itogo_ + 3;
                    }
                    j = itogo_ + 2;
                }
                i = j;
            }
            wb.SaveAs("D:\\reports\\EqDailyReports" + SD.ToString("MMMM", new CultureInfo("en-GB")) + SD.Year + ".xlsx");
           
        }
        static void Main(string[] args)
        {
            GetEquipmentDataDaily(new DateTime(2020,01,01));
        }
    }
}
