package schedule

import (
	"github.com/Luxurioust/excelize"

	"fmt"
	"math"
	"sort"
	"time"

	"github.com/go-steven/sanle/utils"
)

func SaveToExcel(excel string, schedulesData *SchedulesData) error {
	xlsx := excelize.NewFile()

	sheetId, err := SaveStaffStatToExcel(xlsx, schedulesData)
	if err != nil {
		Logger.Error(err)
		return err
	}

	// Set active sheet of the workbook.
	xlsx.SetActiveSheet(sheetId)
	// Save xlsx file by the given path.
	if err := xlsx.SaveAs(excel); err != nil {
		Logger.Error(err)
		return err
	}

	return nil
}

// 保存员工工时结算信息到EXCEL中
func SaveStaffStatToExcel(xlsx *excelize.File, schedulesData *SchedulesData) (int, error) {
	sheet := "工时结算"
	sheetId := xlsx.NewSheet(sheet)

	loc, _ := time.LoadLocation("Asia/Chongqing")

	//Logger.Infof("startDate : %s, endDate: %s", startDate, endDate)
	t_start, err := time.ParseInLocation("2006-01-02", schedulesData.StartDate, loc)
	if err != nil {
		Logger.Error(err)
		return 0, err
	}

	t_end, err := time.ParseInLocation("2006-01-02", schedulesData.EndDate, loc)
	if err != nil {
		Logger.Error(err)
		return 0, err
	}

	dates := []time.Time{}
	t_curr_date := t_start
	startRowId := 0
	for {
		if t_curr_date.Day() == 26 { // 每个月工时结算到25日，26日之后结算到下一个月
			if len(dates) > 0 {
				rowId, err := write_staff_schedule_block(xlsx, sheet, startRowId, schedulesData.Data, dates[0].Format("2006-01-02"), dates[len(dates)-1].Format("2006-01-02"))
				if err != nil {
					return 0, err
				}
				dates = []time.Time{}
				startRowId = rowId + 1
			}
		}
		dates = append(dates, t_curr_date)

		t_curr_date = t_curr_date.Add(24 * time.Hour)
		if t_curr_date.After(t_end) {
			break
		}
	}
	if len(dates) > 0 {
		if _, err := write_staff_schedule_block(xlsx, sheet, startRowId, schedulesData.Data, dates[0].Format("2006-01-02"), dates[len(dates)-1].Format("2006-01-02")); err != nil {
			return 0, err
		}
	}

	return sheetId, nil
}

// 员工工时统计对象
type StaffStat struct {
	StaffName       string  `json:"staff_name" codec:"staff_name,omitempty"`   // 员工姓名
	TotalDayHours   float64 `json:"day_hours" codec:"day_hours,omitempty"`     // 结算白班小时数
	TotalNightHours float64 `json:"night_hours" codec:"night_hours,omitempty"` // 结算夜班小时数
	TotalDays       int     `json:"total_days" codec:"total_days,omitempty"`   // 总上班天数

	YangzhaDayHours   float64 `json:"yangzha_day_hours" codec:"yangzha_day_hours,omitempty"`     // 杨闸白班小时数
	YangzhaNightHours float64 `json:"yangzha_night_hours" codec:"yangzha_night_hours,omitempty"` // 杨闸夜班小时数
	YangzhaDays       int     `json:"yangzha_days" codec:"yangzha_days,omitempty"`               // 杨闸上班天数

	NanceDayHours   float64 `json:"nance_day_hours" codec:"nance_day_hours,omitempty"`     // 南侧白班小时数
	NanceNightHours float64 `json:"nance_night_hours" codec:"nance_night_hours,omitempty"` // 南侧夜班小时数
	NanceDays       int     `json:"nance_days" codec:"nance_days,omitempty"`               // 南侧上班天数

	JiahuiDayHours   float64 `json:"jiahui_day_hours" codec:"jiahui_day_hours,omitempty"`     // 佳汇白班小时数
	JiahuiNightHours float64 `json:"jiahui_night_hours" codec:"jiahui_night_hours,omitempty"` // 佳汇夜班小时数
	JiahuiDays       int     `json:"jiahui_days" codec:"jiahui_days,omitempty"`               // 佳汇上班天数
}

// 以结算月为区块，保存员工月工时结算数据到EXCEL中
// 返回写入的EXCEL结束行ID
func write_staff_schedule_block(xlsx *excelize.File, sheetName string, startRowId int, schedules map[string][]*Schedule, startDate, endDate string) (int, error) {
	loc, _ := time.LoadLocation("Asia/Chongqing")

	Logger.Infof("startDate : %s, endDate: %s", startDate, endDate)
	// 结算月开始日期
	t_start, err := time.ParseInLocation("2006-01-02", startDate, loc)
	if err != nil {
		Logger.Error(err)
		return 0, err
	}

	// 结算月截至日期
	t_end, err := time.ParseInLocation("2006-01-02", endDate, loc)
	if err != nil {
		Logger.Error(err)
		return 0, err
	}

	// 生成考勤日期列表
	dates := []time.Time{}
	t_date := t_start
	for {
		dates = append(dates, t_date)

		t_date = t_date.Add(24 * time.Hour)
		if t_date.After(t_end) {
			break
		}
	}

	// 写入区块标题（汇总字段），占2行， 19列
	startRowId++
	xlsx.SetCellValue(sheetName, fmt.Sprintf("A%d", startRowId), fmt.Sprintf("%d年%d月", t_end.Year(), t_end.Month()))                                     // 结算年月
	xlsx.SetCellValue(sheetName, fmt.Sprintf("B%d", startRowId), fmt.Sprintf("%d/%d-%d/%d", t_start.Month(), t_start.Day(), t_end.Month(), t_end.Day())) // 结算月：开始日期/截至日期
	xlsx.SetCellValue(sheetName, fmt.Sprintf("A%d", startRowId+1), "序号")
	xlsx.SetCellValue(sheetName, fmt.Sprintf("B%d", startRowId+1), "姓名")
	xlsx.SetCellValue(sheetName, fmt.Sprintf("C%d", startRowId+1), "总工时")
	xlsx.SetCellValue(sheetName, fmt.Sprintf("D%d", startRowId+1), "总人天")
	xlsx.SetCellValue(sheetName, fmt.Sprintf("E%d", startRowId+1), "总白班工时")
	xlsx.SetCellValue(sheetName, fmt.Sprintf("F%d", startRowId+1), "总夜班工时")
	xlsx.SetCellValue(sheetName, fmt.Sprintf("G%d", startRowId+1), fmt.Sprintf("%s人天", WORK_SHOP_YANGZHA))
	xlsx.SetCellValue(sheetName, fmt.Sprintf("H%d", startRowId+1), fmt.Sprintf("%s白班工时", WORK_SHOP_YANGZHA))
	xlsx.SetCellValue(sheetName, fmt.Sprintf("I%d", startRowId+1), fmt.Sprintf("%s夜班工时", WORK_SHOP_YANGZHA))
	xlsx.SetCellValue(sheetName, fmt.Sprintf("J%d", startRowId+1), fmt.Sprintf("%s工时合计", WORK_SHOP_YANGZHA))
	xlsx.SetCellValue(sheetName, fmt.Sprintf("K%d", startRowId+1), fmt.Sprintf("%s人天", WORK_SHOP_NANCE))
	xlsx.SetCellValue(sheetName, fmt.Sprintf("L%d", startRowId+1), fmt.Sprintf("%s白班工时", WORK_SHOP_NANCE))
	xlsx.SetCellValue(sheetName, fmt.Sprintf("M%d", startRowId+1), fmt.Sprintf("%s夜班工时", WORK_SHOP_NANCE))
	xlsx.SetCellValue(sheetName, fmt.Sprintf("N%d", startRowId+1), fmt.Sprintf("%s工时合计", WORK_SHOP_NANCE))
	xlsx.SetCellValue(sheetName, fmt.Sprintf("O%d", startRowId+1), fmt.Sprintf("%s人天", WORK_SHOP_JIAHUI))
	xlsx.SetCellValue(sheetName, fmt.Sprintf("P%d", startRowId+1), fmt.Sprintf("%s白班工时", WORK_SHOP_JIAHUI))
	xlsx.SetCellValue(sheetName, fmt.Sprintf("Q%d", startRowId+1), fmt.Sprintf("%s夜班工时", WORK_SHOP_JIAHUI))
	xlsx.SetCellValue(sheetName, fmt.Sprintf("R%d", startRowId+1), fmt.Sprintf("%s工时合计", WORK_SHOP_JIAHUI))
	xlsx.SetCellValue(sheetName, fmt.Sprintf("S%d", startRowId+1), "") // 空1列作为间隔

	columns := utils.GetExcelColumns(26 * 3)
	dateColumns := make(map[string]string) // 保存具体考勤日期对应的EXCEL列名
	// 写入区块标题考勤日期及对应的星期几，占2行
	for idx, v := range dates {
		column := columns[idx+19]                                                                                         // 日期列从第20列开始
		xlsx.SetCellValue(sheetName, fmt.Sprintf("%s%d", column, startRowId), utils.GetWeekDay(v))                        // 考勤日期对应的星期几
		xlsx.SetCellValue(sheetName, fmt.Sprintf("%s%d", column, startRowId+1), fmt.Sprintf("%d/%d", v.Month(), v.Day())) // 考勤日期

		dateColumns[v.Format("2006-01-02")] = column
	}

	// 按员工姓名排序
	staffNames := []string{}
	for k, _ := range schedules {
		staffNames = append(staffNames, k)
	}
	sort.Strings(staffNames)

	summaryStat := &StaffStat{}

	rowId := startRowId + 2 // 跳过区块标题2行
	for _, staffName := range staffNames {
		rows := []*Schedule{}
		for _, v := range schedules[staffName] {
			if v.ScheduleDate < startDate || v.ScheduleDate > endDate { // 只处理给定结算日期内的排班记录
				continue
			}
			rows = append(rows, v)
		}
		if len(rows) == 0 {
			continue
		}

		stat := &StaffStat{
			StaffName: staffName,
			TotalDays: len(rows), // 人天
		}
		for _, v := range rows {
			t_start_time, err := time.ParseInLocation("2006-01-02 15:04", v.StartTime, loc)
			if err != nil {
				Logger.Error(err)
				return 0, err
			}

			t_end_time, err := time.ParseInLocation("2006-01-02 15:04", v.EndTime, loc)
			if err != nil {
				Logger.Error(err)
				return 0, err
			}

			// 排班时间反格式化
			var start_hour_str, end_hour_str string
			if t_start_time.Minute() > 0 {
				start_hour_str = fmt.Sprintf("%d:%02d", t_start_time.Hour(), t_start_time.Minute())
			} else {
				start_hour_str = fmt.Sprintf("%d", t_start_time.Hour())
			}
			if t_end_time.Minute() > 0 {
				end_hour_str = fmt.Sprintf("%d:%02d", t_end_time.Hour(), t_end_time.Minute())
			} else {
				end_hour := t_end_time.Hour()
				if end_hour == 0 {
					end_hour = 24
				}
				end_hour_str = fmt.Sprintf("%d", end_hour)
			}
			stat.TotalDayHours += v.WageDayHours
			stat.TotalNightHours += v.WageNightHours

			switch v.WorkShop {
			case WORK_SHOP_YANGZHA:
				stat.YangzhaDayHours += v.WageDayHours
				stat.YangzhaNightHours += v.WageNightHours
				stat.YangzhaDays++
			case WORK_SHOP_NANCE:
				stat.NanceDayHours += v.WageDayHours
				stat.NanceNightHours += v.WageNightHours
				stat.NanceDays++
			case WORK_SHOP_JIAHUI:
				stat.JiahuiDayHours += v.WageDayHours
				stat.JiahuiNightHours += v.WageNightHours
				stat.JiahuiDays++
			}
			xlsx.SetCellValue(sheetName, fmt.Sprintf("%s%d", dateColumns[v.ScheduleDate], rowId), fmt.Sprintf("%s-%s(%s)", start_hour_str, end_hour_str, string([]rune(v.WorkShop)[0])))
		}

		xlsx.SetCellValue(sheetName, fmt.Sprintf("A%d", rowId), rowId-(startRowId+2)+1)                      // 序号
		xlsx.SetCellValue(sheetName, fmt.Sprintf("B%d", rowId), stat.StaffName)                              // 员工姓名
		xlsx.SetCellValue(sheetName, fmt.Sprintf("C%d", rowId), stat.TotalDayHours+stat.TotalNightHours)     // 总工时
		xlsx.SetCellValue(sheetName, fmt.Sprintf("D%d", rowId), stat.TotalDays)                              // 总人天
		xlsx.SetCellValue(sheetName, fmt.Sprintf("E%d", rowId), stat.TotalDayHours)                          // 总白班工时
		xlsx.SetCellValue(sheetName, fmt.Sprintf("F%d", rowId), stat.TotalNightHours)                        // 总夜班工时
		xlsx.SetCellValue(sheetName, fmt.Sprintf("G%d", rowId), stat.YangzhaDays)                            // 杨闸人天
		xlsx.SetCellValue(sheetName, fmt.Sprintf("H%d", rowId), stat.YangzhaDayHours)                        // 杨闸白班工时
		xlsx.SetCellValue(sheetName, fmt.Sprintf("I%d", rowId), stat.YangzhaNightHours)                      // 杨闸夜班工时
		xlsx.SetCellValue(sheetName, fmt.Sprintf("J%d", rowId), stat.YangzhaDayHours+stat.YangzhaNightHours) // 杨闸工时合计
		xlsx.SetCellValue(sheetName, fmt.Sprintf("K%d", rowId), stat.NanceDays)                              // 南侧人天
		xlsx.SetCellValue(sheetName, fmt.Sprintf("L%d", rowId), stat.NanceDayHours)                          // 南侧白班工时
		xlsx.SetCellValue(sheetName, fmt.Sprintf("M%d", rowId), stat.NanceNightHours)                        // 南侧夜班工时
		xlsx.SetCellValue(sheetName, fmt.Sprintf("N%d", rowId), stat.NanceDayHours+stat.NanceNightHours)     // 南侧工时合计
		xlsx.SetCellValue(sheetName, fmt.Sprintf("O%d", rowId), stat.JiahuiDays)                             // 佳汇人天
		xlsx.SetCellValue(sheetName, fmt.Sprintf("P%d", rowId), stat.JiahuiDayHours)                         // 佳汇白班工时
		xlsx.SetCellValue(sheetName, fmt.Sprintf("Q%d", rowId), stat.JiahuiNightHours)                       // 佳汇夜班工时
		xlsx.SetCellValue(sheetName, fmt.Sprintf("R%d", rowId), stat.JiahuiDayHours+stat.JiahuiNightHours)   // 佳汇工时合计

		summaryStat.TotalDays += stat.TotalDays                 // 总人天合计
		summaryStat.TotalDayHours += stat.TotalDayHours         // 总白班工时合计
		summaryStat.TotalNightHours += stat.TotalNightHours     // 总夜班工时合计
		summaryStat.YangzhaDays += stat.YangzhaDays             // 杨闸人天合计
		summaryStat.YangzhaDayHours += stat.YangzhaDayHours     // 杨闸白班工时合计
		summaryStat.YangzhaNightHours += stat.YangzhaNightHours // 杨闸夜班工时合计
		summaryStat.NanceDays += stat.NanceDays                 // 南侧人天合计
		summaryStat.NanceDayHours += stat.NanceDayHours         // 南侧白班工时合计
		summaryStat.NanceNightHours += stat.NanceNightHours     // 南侧夜班工时合计
		summaryStat.JiahuiDays += stat.JiahuiDays               // 佳汇人天合计
		summaryStat.JiahuiDayHours += stat.JiahuiDayHours       // 佳汇白班工时合计
		summaryStat.JiahuiNightHours += stat.JiahuiNightHours   // 佳汇夜班工时合计

		rowId++
	}

	xlsx.SetCellValue(sheetName, fmt.Sprintf("A%d", rowId), "合计")                                                      // 合计标题
	xlsx.SetCellValue(sheetName, fmt.Sprintf("B%d", rowId), fmt.Sprintf("%d天", len(dates)))                            // 结算月天数
	xlsx.SetCellValue(sheetName, fmt.Sprintf("C%d", rowId), summaryStat.TotalDayHours+summaryStat.TotalNightHours)     // 总工时合计
	xlsx.SetCellValue(sheetName, fmt.Sprintf("D%d", rowId), summaryStat.TotalDays)                                     // 总人天合计
	xlsx.SetCellValue(sheetName, fmt.Sprintf("E%d", rowId), summaryStat.TotalDayHours)                                 // 总白班工时合计
	xlsx.SetCellValue(sheetName, fmt.Sprintf("F%d", rowId), summaryStat.TotalNightHours)                               // 总夜班工时合计
	xlsx.SetCellValue(sheetName, fmt.Sprintf("G%d", rowId), summaryStat.YangzhaDays)                                   // 杨闸人天合计
	xlsx.SetCellValue(sheetName, fmt.Sprintf("H%d", rowId), summaryStat.YangzhaDayHours)                               // 杨闸白班工时合计
	xlsx.SetCellValue(sheetName, fmt.Sprintf("I%d", rowId), summaryStat.YangzhaNightHours)                             // 杨闸夜班工时合计
	xlsx.SetCellValue(sheetName, fmt.Sprintf("J%d", rowId), summaryStat.YangzhaDayHours+summaryStat.YangzhaNightHours) // 杨闸工时合计
	xlsx.SetCellValue(sheetName, fmt.Sprintf("K%d", rowId), summaryStat.NanceDays)                                     // 南侧人天合计
	xlsx.SetCellValue(sheetName, fmt.Sprintf("L%d", rowId), summaryStat.NanceDayHours)                                 // 南侧白班工时合计
	xlsx.SetCellValue(sheetName, fmt.Sprintf("M%d", rowId), summaryStat.NanceNightHours)                               // 南侧夜班工时合计
	xlsx.SetCellValue(sheetName, fmt.Sprintf("N%d", rowId), summaryStat.NanceDayHours+summaryStat.NanceNightHours)     // 南侧工时合计
	xlsx.SetCellValue(sheetName, fmt.Sprintf("O%d", rowId), summaryStat.JiahuiDays)                                    // 佳汇人天合计
	xlsx.SetCellValue(sheetName, fmt.Sprintf("P%d", rowId), summaryStat.JiahuiDayHours)                                // 佳汇白班工时合计
	xlsx.SetCellValue(sheetName, fmt.Sprintf("Q%d", rowId), summaryStat.JiahuiNightHours)                              // 佳汇夜班工时合计
	xlsx.SetCellValue(sheetName, fmt.Sprintf("R%d", rowId), summaryStat.JiahuiDayHours+summaryStat.JiahuiNightHours)   // 佳汇工时合计
	rowId++

	monDays := float64(len(dates))
	xlsx.SetCellValue(sheetName, fmt.Sprintf("A%d", rowId), "合计/天")
	xlsx.SetCellValue(sheetName, fmt.Sprintf("B%d", rowId), "")
	xlsx.SetCellValue(sheetName, fmt.Sprintf("C%d", rowId), math.Round(100.0*(summaryStat.TotalDayHours+summaryStat.TotalNightHours)/monDays)/100.0)     // 总平均工时/天
	xlsx.SetCellValue(sheetName, fmt.Sprintf("D%d", rowId), math.Round(100.0*float64(summaryStat.TotalDays)/monDays)/100.0)                              // 总平均人天/天
	xlsx.SetCellValue(sheetName, fmt.Sprintf("E%d", rowId), math.Round(100.0*summaryStat.TotalDayHours/monDays)/100.0)                                   // 总平均白班工时/天
	xlsx.SetCellValue(sheetName, fmt.Sprintf("F%d", rowId), math.Round(100.0*summaryStat.TotalNightHours/monDays)/100.0)                                 // 总平均夜班工时/天
	xlsx.SetCellValue(sheetName, fmt.Sprintf("G%d", rowId), math.Round(100.0*float64(summaryStat.YangzhaDays)/monDays)/100.0)                            // 杨闸平均人天/天
	xlsx.SetCellValue(sheetName, fmt.Sprintf("H%d", rowId), math.Round(100.0*summaryStat.YangzhaDayHours/monDays)/100.0)                                 // 杨闸平均白班工时/天
	xlsx.SetCellValue(sheetName, fmt.Sprintf("I%d", rowId), math.Round(100.0*summaryStat.YangzhaNightHours/monDays)/100.0)                               // 杨闸平均夜班工时/天
	xlsx.SetCellValue(sheetName, fmt.Sprintf("J%d", rowId), math.Round(100.0*(summaryStat.YangzhaDayHours+summaryStat.YangzhaNightHours)/monDays)/100.0) // 杨闸平均工时/天
	xlsx.SetCellValue(sheetName, fmt.Sprintf("K%d", rowId), math.Round(100.0*float64(summaryStat.NanceDays)/monDays)/100.0)                              // 南侧平均人天/天
	xlsx.SetCellValue(sheetName, fmt.Sprintf("L%d", rowId), math.Round(100.0*summaryStat.NanceDayHours/monDays)/100.0)                                   // 南侧平均白班工时/天
	xlsx.SetCellValue(sheetName, fmt.Sprintf("M%d", rowId), math.Round(100.0*summaryStat.NanceNightHours/monDays)/100.0)                                 // 南侧平均夜班工时/天
	xlsx.SetCellValue(sheetName, fmt.Sprintf("N%d", rowId), math.Round(100.0*(summaryStat.NanceDayHours+summaryStat.NanceNightHours)/monDays)/100.0)     // 南侧平均工时/天
	xlsx.SetCellValue(sheetName, fmt.Sprintf("O%d", rowId), math.Round(100.0*float64(summaryStat.JiahuiDays)/monDays)/100.0)                             // 佳汇平均人天/天
	xlsx.SetCellValue(sheetName, fmt.Sprintf("P%d", rowId), math.Round(100.0*summaryStat.JiahuiDayHours/monDays)/100.0)                                  // 佳汇平均白班工时/天
	xlsx.SetCellValue(sheetName, fmt.Sprintf("Q%d", rowId), math.Round(100.0*summaryStat.JiahuiNightHours/monDays)/100.0)                                // 佳汇平均夜班工时/天
	xlsx.SetCellValue(sheetName, fmt.Sprintf("R%d", rowId), math.Round(100.0*(summaryStat.JiahuiDayHours+summaryStat.JiahuiNightHours)/monDays)/100.0)   // 佳汇平均工时/天
	rowId++

	return rowId, nil
}
