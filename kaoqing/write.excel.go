package kaoqing

import (
	"github.com/Luxurioust/excelize"

	"fmt"
	"sort"
	"time"

	"github.com/go-steven/sanle/schedule"
)

func SaveToExcel(excel string, origRecords map[string][]*OrigRecord, kaoqings map[string][]*KaoQing, kaoqingAggregate map[string]*KaoQingAggregate, startDate, endDate string, filterName string) error {
	xlsx := excelize.NewFile()

	if _, err := SaveOrignRecordsToExcel(xlsx, origRecords, startDate, endDate, filterName); err != nil {
		Logger.Error(err)
		return err
	}
	if _, err := SaveKaoqingToExcel(xlsx, kaoqings, startDate, endDate, filterName); err != nil {
		Logger.Error(err)
		return err
	}

	if _, err := SaveInvalidKaoqingToExcel(xlsx, kaoqings, startDate, endDate, filterName); err != nil {
		Logger.Error(err)
		return err
	}

	sheetId, err := SaveKaoqingAggregateToExcel(xlsx, kaoqingAggregate, startDate, endDate, filterName)
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

func SaveOrignRecordsToExcel(xlsx *excelize.File, origRecords map[string][]*OrigRecord, startDate, endDate string, filterName string) (int, error) {
	data := []*OrigRecord{}
	for _, records := range origRecords {
		for _, v := range records {
			data = append(data, v)
		}
	}
	sort.Slice(data, func(i, j int) bool {
		if data[i].StaffName == data[j].StaffName {
			return data[i].RecordOn < data[j].RecordOn
		} else {
			return data[i].StaffName < data[j].StaffName
		}
	})

	loc, _ := time.LoadLocation("Asia/Chongqing")
	t_start, err := time.ParseInLocation("2006-01-02", startDate, loc)
	if err != nil {
		Logger.Error(err)
		return 0, err
	}
	startDate = t_start.Add((-1) * 24 * time.Hour).Format("2006-01-02")

	t_end, err := time.ParseInLocation("2006-01-02", endDate, loc)
	if err != nil {
		Logger.Error(err)
		return 0, err
	}
	today, err := time.ParseInLocation("2006-01-02", time.Now().Format("2006-01-02"), loc)
	if err != nil {
		Logger.Error(err)
		return 0, err
	}
	if t_end.After(today) {
		t_end = today.Add((-1) * 24 * time.Hour) // 取昨天
	}
	endDate = t_end.Add(24 * time.Hour).Format("2006-01-02")

	sheetName := "打卡明细"
	sheetId := xlsx.NewSheet(sheetName)

	xlsx.SetCellValue(sheetName, "A1", fmt.Sprintf("%d年%d月", t_end.Year(), t_end.Month()))                                     // 结算年月
	xlsx.SetCellValue(sheetName, "B1", fmt.Sprintf("%d/%d-%d/%d", t_start.Month(), t_start.Day(), t_end.Month(), t_end.Day())) // 结算月：开始日期/截至日期
	xlsx.SetCellValue(sheetName, "C1", fmt.Sprintf("共%d天", int(t_end.Sub(t_start).Hours()/24)))                                // 结算天数

	xlsx.SetCellValue(sheetName, "A2", "姓名")
	xlsx.SetCellValue(sheetName, "B2", "打卡日期")
	xlsx.SetCellValue(sheetName, "C2", "打卡时间")
	xlsx.SetCellValue(sheetName, "D2", "打卡地址")

	// Create a new sheet.
	rowId := 3
	var staffName string
	for _, v := range data {
		if startDate != "" && v.RecordDate < startDate {
			continue
		}
		if endDate != "" && v.RecordDate > endDate {
			continue
		}

		if filterName != "" && v.StaffName != filterName {
			continue
		}

		if staffName != "" && v.StaffName != staffName {
			rowId++
		}

		// Set value of a cell.
		xlsx.SetCellValue(sheetName, fmt.Sprintf("A%d", rowId), v.StaffName)
		xlsx.SetCellValue(sheetName, fmt.Sprintf("B%d", rowId), v.RecordDate)
		xlsx.SetCellValue(sheetName, fmt.Sprintf("C%d", rowId), v.RecordOn)
		xlsx.SetCellValue(sheetName, fmt.Sprintf("D%d", rowId), v.WorkShop)

		staffName = v.StaffName
		rowId++
	}

	return sheetId, nil
}

func SaveKaoqingToExcel(xlsx *excelize.File, records map[string][]*KaoQing, startDate, endDate string, filterName string) (int, error) {
	data := []*KaoQing{}
	for _, val := range records {
		for _, v := range val {
			data = append(data, v)
		}
	}
	sort.Slice(data, func(i, j int) bool {
		if data[i].StaffName == data[j].StaffName {
			return data[i].StartTime < data[j].StartTime
		} else {
			return data[i].StaffName < data[j].StaffName
		}
	})

	loc, _ := time.LoadLocation("Asia/Chongqing")

	t_start, err := time.ParseInLocation("2006-01-02", startDate, loc)
	if err != nil {
		Logger.Error(err)
		return 0, err
	}

	t_end, err := time.ParseInLocation("2006-01-02", endDate, loc)
	if err != nil {
		Logger.Error(err)
		return 0, err
	}

	today, err := time.ParseInLocation("2006-01-02", time.Now().Format("2006-01-02"), loc)
	if err != nil {
		Logger.Error(err)
		return 0, err
	}
	if t_end.After(today) {
		t_end = today.Add((-1) * 24 * time.Hour) // 取昨天
	}

	sheetName := "考勤明细"
	sheetId := xlsx.NewSheet(sheetName)
	xlsx.SetCellValue(sheetName, "A1", fmt.Sprintf("%d年%d月", t_end.Year(), t_end.Month()))                                     // 结算年月
	xlsx.SetCellValue(sheetName, "B1", fmt.Sprintf("%d/%d-%d/%d", t_start.Month(), t_start.Day(), t_end.Month(), t_end.Day())) // 结算月：开始日期/截至日期
	xlsx.SetCellValue(sheetName, "C1", fmt.Sprintf("共%d天", int(t_end.Sub(t_start).Hours()/24)))                                // 结算天数

	xlsx.SetCellValue(sheetName, "A2", "姓名")
	xlsx.SetCellValue(sheetName, "B2", "考勤日期")
	xlsx.SetCellValue(sheetName, "C2", "班次")
	xlsx.SetCellValue(sheetName, "D2", "开始时间")
	xlsx.SetCellValue(sheetName, "E2", "结束时间")
	xlsx.SetCellValue(sheetName, "F2", "上班地点")
	xlsx.SetCellValue(sheetName, "G2", "上班时间")
	xlsx.SetCellValue(sheetName, "H2", "下班时间")
	xlsx.SetCellValue(sheetName, "I2", "打卡状态")
	xlsx.SetCellValue(sheetName, "J2", "迟到分钟数")
	xlsx.SetCellValue(sheetName, "K2", "早退分钟数")
	xlsx.SetCellValue(sheetName, "L2", "加班分钟数")
	xlsx.SetCellValue(sheetName, "M2", "考勤状态")

	rankDesc := map[uint8]string{
		schedule.RANK_DAY:   "白班",
		schedule.RANK_NIGHT: "夜班",
	}
	recordStsDesc := map[uint8]string{
		RECORD_STS_NORMAL:  "正常",
		RECORD_STS_NONE:    "无打卡记录",
		RECORD_STS_NOSTART: "上班未打卡",
		RECORD_STS_NOEND:   "下班未打卡",
	}
	stsDesc := map[uint8]string{
		KAOQING_STS_NORMAL:         "正常",
		KAOQING_STS_LATE:           "迟到",
		KAOQING_STS_EARLY:          "早退",
		KAOQING_STS_LATE_AND_EARLY: "迟到+早退",
		KAOQING_STS_NONE:           "未打卡或者打卡异常",
	}

	// Create a new sheet.
	rowId := 3
	var staffName string
	for _, v := range data {
		if startDate != "" && v.ScheduleDate < startDate {
			continue
		}
		if endDate != "" && v.ScheduleDate > endDate {
			continue
		}

		if filterName != "" && v.StaffName != filterName {
			continue
		}

		if v.ScheduleDate >= today.Format("2006-01-02") {
			continue
		}
		if staffName != "" && v.StaffName != staffName {
			rowId++
		}

		xlsx.SetCellValue(sheetName, fmt.Sprintf("A%d", rowId), v.StaffName)
		xlsx.SetCellValue(sheetName, fmt.Sprintf("B%d", rowId), v.ScheduleDate)
		xlsx.SetCellValue(sheetName, fmt.Sprintf("C%d", rowId), rankDesc[v.Rank])
		xlsx.SetCellValue(sheetName, fmt.Sprintf("D%d", rowId), v.StartTime)
		xlsx.SetCellValue(sheetName, fmt.Sprintf("E%d", rowId), v.EndTime)
		xlsx.SetCellValue(sheetName, fmt.Sprintf("F%d", rowId), v.WorkShop)
		if v.StartRecord != nil {
			xlsx.SetCellValue(sheetName, fmt.Sprintf("G%d", rowId), v.StartRecord.RecordOn)
		} else {
			xlsx.SetCellValue(sheetName, fmt.Sprintf("G%d", rowId), "")
		}
		if v.EndRecord != nil {
			xlsx.SetCellValue(sheetName, fmt.Sprintf("H%d", rowId), v.EndRecord.RecordOn)
		} else {
			xlsx.SetCellValue(sheetName, fmt.Sprintf("H%d", rowId), "")
		}

		xlsx.SetCellValue(sheetName, fmt.Sprintf("I%d", rowId), recordStsDesc[v.RecordSts])
		if v.LateMinutes > 0 {
			xlsx.SetCellValue(sheetName, fmt.Sprintf("J%d", rowId), v.LateMinutes)
		} else {
			xlsx.SetCellValue(sheetName, fmt.Sprintf("J%d", rowId), "")
		}
		if v.EarlyMinutes > 0 {
			xlsx.SetCellValue(sheetName, fmt.Sprintf("K%d", rowId), v.EarlyMinutes)
		} else {
			xlsx.SetCellValue(sheetName, fmt.Sprintf("K%d", rowId), "")
		}
		if v.OverMinutes > 0 {
			xlsx.SetCellValue(sheetName, fmt.Sprintf("L%d", rowId), v.OverMinutes)
		} else {
			xlsx.SetCellValue(sheetName, fmt.Sprintf("L%d", rowId), "")
		}
		xlsx.SetCellValue(sheetName, fmt.Sprintf("M%d", rowId), stsDesc[v.Sts])

		staffName = v.StaffName
		rowId++
	}

	return sheetId, nil
}

func SaveInvalidKaoqingToExcel(xlsx *excelize.File, records map[string][]*KaoQing, startDate, endDate string, filterName string) (int, error) {
	data := []*KaoQing{}
	for _, val := range records {
		for _, v := range val {
			data = append(data, v)
		}
	}
	sort.Slice(data, func(i, j int) bool {
		if data[i].StaffName == data[j].StaffName {
			return data[i].StartTime < data[j].StartTime
		} else {
			return data[i].StaffName < data[j].StaffName
		}
	})

	loc, _ := time.LoadLocation("Asia/Chongqing")

	t_start, err := time.ParseInLocation("2006-01-02", startDate, loc)
	if err != nil {
		Logger.Error(err)
		return 0, err
	}

	t_end, err := time.ParseInLocation("2006-01-02", endDate, loc)
	if err != nil {
		Logger.Error(err)
		return 0, err
	}

	today, err := time.ParseInLocation("2006-01-02", time.Now().Format("2006-01-02"), loc)
	if err != nil {
		Logger.Error(err)
		return 0, err
	}
	if t_end.After(today) {
		t_end = today.Add((-1) * 24 * time.Hour) // 取昨天
	}

	sheetName := "异常考勤"
	sheetId := xlsx.NewSheet(sheetName)
	xlsx.SetCellValue(sheetName, "A1", fmt.Sprintf("%d年%d月", t_end.Year(), t_end.Month()))                                     // 结算年月
	xlsx.SetCellValue(sheetName, "B1", fmt.Sprintf("%d/%d-%d/%d", t_start.Month(), t_start.Day(), t_end.Month(), t_end.Day())) // 结算月：开始日期/截至日期
	xlsx.SetCellValue(sheetName, "C1", fmt.Sprintf("共%d天", int(t_end.Sub(t_start).Hours()/24)))                                // 结算天数

	xlsx.SetCellValue(sheetName, "A2", "姓名")
	xlsx.SetCellValue(sheetName, "B2", "考勤日期")
	xlsx.SetCellValue(sheetName, "C2", "班次")
	xlsx.SetCellValue(sheetName, "D2", "开始时间")
	xlsx.SetCellValue(sheetName, "E2", "结束时间")
	xlsx.SetCellValue(sheetName, "F2", "上班地点")
	xlsx.SetCellValue(sheetName, "G2", "上班时间")
	xlsx.SetCellValue(sheetName, "H2", "下班时间")
	xlsx.SetCellValue(sheetName, "I2", "打卡状态")
	xlsx.SetCellValue(sheetName, "J2", "迟到分钟数")
	xlsx.SetCellValue(sheetName, "K2", "早退分钟数")
	xlsx.SetCellValue(sheetName, "L2", "加班分钟数")
	xlsx.SetCellValue(sheetName, "M2", "考勤状态")

	rankDesc := map[uint8]string{
		schedule.RANK_DAY:   "白班",
		schedule.RANK_NIGHT: "夜班",
	}
	recordStsDesc := map[uint8]string{
		RECORD_STS_NORMAL:  "正常",
		RECORD_STS_NONE:    "无打卡记录",
		RECORD_STS_NOSTART: "上班未打卡",
		RECORD_STS_NOEND:   "下班未打卡",
	}
	stsDesc := map[uint8]string{
		KAOQING_STS_NORMAL:         "正常",
		KAOQING_STS_LATE:           "迟到",
		KAOQING_STS_EARLY:          "早退",
		KAOQING_STS_LATE_AND_EARLY: "迟到+早退",
		KAOQING_STS_NONE:           "未打卡或者打卡异常",
	}

	// Create a new sheet.
	rowId := 3
	var staffName string
	for _, v := range data {
		if v.Sts == KAOQING_STS_NORMAL {
			continue
		}

		if startDate != "" && v.ScheduleDate < startDate {
			continue
		}
		if endDate != "" && v.ScheduleDate > endDate {
			continue
		}

		if filterName != "" && v.StaffName != filterName {
			continue
		}

		if v.ScheduleDate >= today.Format("2006-01-02") {
			continue
		}
		if staffName != "" && v.StaffName != staffName {
			rowId++
		}

		xlsx.SetCellValue(sheetName, fmt.Sprintf("A%d", rowId), v.StaffName)
		xlsx.SetCellValue(sheetName, fmt.Sprintf("B%d", rowId), v.ScheduleDate)
		xlsx.SetCellValue(sheetName, fmt.Sprintf("C%d", rowId), rankDesc[v.Rank])
		xlsx.SetCellValue(sheetName, fmt.Sprintf("D%d", rowId), v.StartTime)
		xlsx.SetCellValue(sheetName, fmt.Sprintf("E%d", rowId), v.EndTime)
		xlsx.SetCellValue(sheetName, fmt.Sprintf("F%d", rowId), v.WorkShop)
		if v.StartRecord != nil {
			xlsx.SetCellValue(sheetName, fmt.Sprintf("G%d", rowId), v.StartRecord.RecordOn)
		} else {
			xlsx.SetCellValue(sheetName, fmt.Sprintf("G%d", rowId), "")
		}
		if v.EndRecord != nil {
			xlsx.SetCellValue(sheetName, fmt.Sprintf("H%d", rowId), v.EndRecord.RecordOn)
		} else {
			xlsx.SetCellValue(sheetName, fmt.Sprintf("H%d", rowId), "")
		}

		xlsx.SetCellValue(sheetName, fmt.Sprintf("I%d", rowId), recordStsDesc[v.RecordSts])
		if v.LateMinutes > 0 {
			xlsx.SetCellValue(sheetName, fmt.Sprintf("J%d", rowId), v.LateMinutes)
		} else {
			xlsx.SetCellValue(sheetName, fmt.Sprintf("J%d", rowId), "")
		}
		if v.EarlyMinutes > 0 {
			xlsx.SetCellValue(sheetName, fmt.Sprintf("K%d", rowId), v.EarlyMinutes)
		} else {
			xlsx.SetCellValue(sheetName, fmt.Sprintf("K%d", rowId), "")
		}
		if v.OverMinutes > 0 {
			xlsx.SetCellValue(sheetName, fmt.Sprintf("L%d", rowId), v.OverMinutes)
		} else {
			xlsx.SetCellValue(sheetName, fmt.Sprintf("L%d", rowId), "")
		}
		xlsx.SetCellValue(sheetName, fmt.Sprintf("M%d", rowId), stsDesc[v.Sts])

		staffName = v.StaffName
		rowId++
	}

	return sheetId, nil
}

func SaveKaoqingAggregateToExcel(xlsx *excelize.File, records map[string]*KaoQingAggregate, startDate, endDate string, filterName string) (int, error) {
	loc, _ := time.LoadLocation("Asia/Chongqing")

	t_start, err := time.ParseInLocation("2006-01-02", startDate, loc)
	if err != nil {
		Logger.Error(err)
		return 0, err
	}

	t_end, err := time.ParseInLocation("2006-01-02", endDate, loc)
	if err != nil {
		Logger.Error(err)
		return 0, err
	}

	today, err := time.ParseInLocation("2006-01-02", time.Now().Format("2006-01-02"), loc)
	if err != nil {
		Logger.Error(err)
		return 0, err
	}
	if t_end.After(today) {
		t_end = today.Add((-1) * 24 * time.Hour) // 取昨天
	}

	// 按员工姓名排序
	staffNames := []string{}
	for k, _ := range records {
		staffNames = append(staffNames, k)
	}
	sort.Strings(staffNames)

	sheetName := "考勤汇总"
	sheetId := xlsx.NewSheet(sheetName)

	xlsx.SetCellValue(sheetName, "A1", fmt.Sprintf("%d年%d月", t_end.Year(), t_end.Month()))                                     // 结算年月
	xlsx.SetCellValue(sheetName, "B1", fmt.Sprintf("%d/%d-%d/%d", t_start.Month(), t_start.Day(), t_end.Month(), t_end.Day())) // 结算月：开始日期/截至日期
	xlsx.SetCellValue(sheetName, "C1", fmt.Sprintf("共%d天", int(t_end.Sub(t_start).Hours()/24)))                                // 结算天数

	xlsx.SetCellValue(sheetName, "A2", "序号")
	xlsx.SetCellValue(sheetName, "B2", "姓名")
	xlsx.SetCellValue(sheetName, "C2", "总人天")
	xlsx.SetCellValue(sheetName, "D2", "考勤正常")
	xlsx.SetCellValue(sheetName, "E2", "迟到")
	xlsx.SetCellValue(sheetName, "F2", "漏打卡")
	xlsx.SetCellValue(sheetName, "G2", "未打卡")
	xlsx.SetCellValue(sheetName, "H2", "早退")
	xlsx.SetCellValue(sheetName, "I2", "有加班")
	xlsx.SetCellValue(sheetName, "J2", "打卡正常")
	xlsx.SetCellValue(sheetName, "K2", "佳汇")
	xlsx.SetCellValue(sheetName, "L2", "杨闸")
	xlsx.SetCellValue(sheetName, "M2", "南侧")

	daysSummary := &KaoQingAggregate{}
	staffCntSummary := &KaoQingAggregate{}

	rowId := 3
	for _, staffName := range staffNames {
		v := records[staffName]
		if filterName != "" && v.StaffName != filterName {
			continue
		}
		if v.TotalDays == 0 {
			continue
		}

		staffCntSummary.TotalDays++

		xlsx.SetCellValue(sheetName, fmt.Sprintf("A%d", rowId), rowId-3+1)   // 序号
		xlsx.SetCellValue(sheetName, fmt.Sprintf("B%d", rowId), v.StaffName) // 姓名
		xlsx.SetCellValue(sheetName, fmt.Sprintf("C%d", rowId), v.TotalDays) // 总人天
		if v.NormalDays > 0 {
			xlsx.SetCellValue(sheetName, fmt.Sprintf("D%d", rowId), v.NormalDays) // 考勤正常
		}
		if v.NormalDays > 0 && v.NormalDays == v.TotalDays {
			staffCntSummary.NormalDays++
		}
		if v.LateDays > 0 {
			xlsx.SetCellValue(sheetName, fmt.Sprintf("E%d", rowId), v.LateDays) // 迟到天数
			staffCntSummary.LateDays++
		}
		if v.RecordNoStartDays > 0 {
			staffCntSummary.RecordNoStartDays++
		}
		if v.RecordNoEndDays > 0 {
			staffCntSummary.RecordNoEndDays++
		}
		if v.RecordNoStartDays+v.RecordNoEndDays > 0 {
			xlsx.SetCellValue(sheetName, fmt.Sprintf("F%d", rowId), v.RecordNoStartDays+v.RecordNoEndDays) // 漏打卡天数
		}
		if v.RecordNoneDays > 0 {
			xlsx.SetCellValue(sheetName, fmt.Sprintf("G%d", rowId), v.RecordNoneDays) // 全部未打卡天数
			staffCntSummary.RecordNoneDays++
		}
		if v.EarlyDays > 0 {
			xlsx.SetCellValue(sheetName, fmt.Sprintf("H%d", rowId), v.EarlyDays) // 早退天数
			staffCntSummary.EarlyDays++
		}
		if v.OverDays > 0 {
			xlsx.SetCellValue(sheetName, fmt.Sprintf("I%d", rowId), v.OverDays) // 有加班天数
			staffCntSummary.OverDays++
		}
		if v.RecordNormalDays > 0 {
			xlsx.SetCellValue(sheetName, fmt.Sprintf("J%d", rowId), v.RecordNormalDays) // 打卡正常天数
		}
		if v.RecordNormalDays > 0 && v.RecordNormalDays == v.TotalDays {
			staffCntSummary.RecordNormalDays++
		}
		if v.JiahuiDays > 0 {
			xlsx.SetCellValue(sheetName, fmt.Sprintf("K%d", rowId), v.JiahuiDays) // 佳汇天数
			staffCntSummary.JiahuiDays++
		}
		if v.YangzhaDays > 0 {
			xlsx.SetCellValue(sheetName, fmt.Sprintf("L%d", rowId), v.YangzhaDays) // 杨闸天数
			staffCntSummary.YangzhaDays++
		}
		if v.NanceDays > 0 {
			xlsx.SetCellValue(sheetName, fmt.Sprintf("M%d", rowId), v.NanceDays) // 南侧天数
			staffCntSummary.NanceDays++
		}

		daysSummary.TotalDays += v.TotalDays
		daysSummary.YangzhaDays += v.YangzhaDays
		daysSummary.NanceDays += v.NanceDays
		daysSummary.JiahuiDays += v.JiahuiDays
		daysSummary.NormalDays += v.NormalDays
		daysSummary.LateDays += v.LateDays
		daysSummary.EarlyDays += v.EarlyDays
		daysSummary.OverDays += v.OverDays
		daysSummary.RecordNormalDays += v.RecordNormalDays
		daysSummary.RecordNoneDays += v.RecordNoneDays
		daysSummary.RecordNoStartDays += v.RecordNoStartDays
		daysSummary.RecordNoEndDays += v.RecordNoEndDays

		rowId++
	}

	// 写入人天合计
	xlsx.SetCellValue(sheetName, fmt.Sprintf("A%d", rowId), "人天合计")
	xlsx.SetCellValue(sheetName, fmt.Sprintf("B%d", rowId), "")                    // 姓名
	xlsx.SetCellValue(sheetName, fmt.Sprintf("C%d", rowId), daysSummary.TotalDays) // 总人天
	if daysSummary.NormalDays > 0 {
		xlsx.SetCellValue(sheetName, fmt.Sprintf("D%d", rowId), daysSummary.NormalDays) // 考勤正常
	}
	if daysSummary.LateDays > 0 {
		xlsx.SetCellValue(sheetName, fmt.Sprintf("E%d", rowId), daysSummary.LateDays) // 迟到天数
	}
	if daysSummary.RecordNoStartDays+daysSummary.RecordNoEndDays > 0 {
		xlsx.SetCellValue(sheetName, fmt.Sprintf("F%d", rowId), daysSummary.RecordNoStartDays+daysSummary.RecordNoEndDays) // 漏打卡天数
	}
	if daysSummary.RecordNoneDays > 0 {
		xlsx.SetCellValue(sheetName, fmt.Sprintf("G%d", rowId), daysSummary.RecordNoneDays) // 全部未打卡天数
	}
	if daysSummary.EarlyDays > 0 {
		xlsx.SetCellValue(sheetName, fmt.Sprintf("H%d", rowId), daysSummary.EarlyDays) // 早退天数
	}
	if daysSummary.OverDays > 0 {
		xlsx.SetCellValue(sheetName, fmt.Sprintf("I%d", rowId), daysSummary.OverDays) // 有加班天数
	}
	if daysSummary.RecordNormalDays > 0 {
		xlsx.SetCellValue(sheetName, fmt.Sprintf("J%d", rowId), daysSummary.RecordNormalDays) // 打卡正常天数
	}
	if daysSummary.JiahuiDays > 0 {
		xlsx.SetCellValue(sheetName, fmt.Sprintf("K%d", rowId), daysSummary.JiahuiDays) // 佳汇天数
	}
	if daysSummary.YangzhaDays > 0 {
		xlsx.SetCellValue(sheetName, fmt.Sprintf("L%d", rowId), daysSummary.YangzhaDays) // 杨闸天数
	}
	if daysSummary.NanceDays > 0 {
		xlsx.SetCellValue(sheetName, fmt.Sprintf("M%d", rowId), daysSummary.NanceDays) // 南侧天数
	}
	rowId++

	// 写入相关人数合计
	xlsx.SetCellValue(sheetName, fmt.Sprintf("A%d", rowId), "相关人数")
	xlsx.SetCellValue(sheetName, fmt.Sprintf("B%d", rowId), "")                        // 姓名
	xlsx.SetCellValue(sheetName, fmt.Sprintf("C%d", rowId), staffCntSummary.TotalDays) // 总人天
	if staffCntSummary.NormalDays > 0 {
		xlsx.SetCellValue(sheetName, fmt.Sprintf("D%d", rowId), staffCntSummary.NormalDays) // 考勤正常
	}
	if staffCntSummary.LateDays > 0 {
		xlsx.SetCellValue(sheetName, fmt.Sprintf("E%d", rowId), staffCntSummary.LateDays) // 迟到天数
	}
	if staffCntSummary.RecordNoStartDays+staffCntSummary.RecordNoEndDays > 0 {
		xlsx.SetCellValue(sheetName, fmt.Sprintf("F%d", rowId), staffCntSummary.RecordNoStartDays+staffCntSummary.RecordNoEndDays) // 漏打卡天数
	}
	if staffCntSummary.RecordNoneDays > 0 {
		xlsx.SetCellValue(sheetName, fmt.Sprintf("G%d", rowId), staffCntSummary.RecordNoneDays) // 全部未打卡天数
	}
	if staffCntSummary.EarlyDays > 0 {
		xlsx.SetCellValue(sheetName, fmt.Sprintf("H%d", rowId), staffCntSummary.EarlyDays) // 早退天数
	}
	if staffCntSummary.OverDays > 0 {
		xlsx.SetCellValue(sheetName, fmt.Sprintf("I%d", rowId), staffCntSummary.OverDays) // 有加班天数
	}
	if staffCntSummary.RecordNormalDays > 0 {
		xlsx.SetCellValue(sheetName, fmt.Sprintf("J%d", rowId), staffCntSummary.RecordNormalDays) // 打卡正常天数
	}
	if staffCntSummary.JiahuiDays > 0 {
		xlsx.SetCellValue(sheetName, fmt.Sprintf("K%d", rowId), staffCntSummary.JiahuiDays) // 佳汇天数
	}
	if staffCntSummary.YangzhaDays > 0 {
		xlsx.SetCellValue(sheetName, fmt.Sprintf("L%d", rowId), staffCntSummary.YangzhaDays) // 杨闸天数
	}
	if staffCntSummary.NanceDays > 0 {
		xlsx.SetCellValue(sheetName, fmt.Sprintf("M%d", rowId), staffCntSummary.NanceDays) // 南侧天数
	}
	rowId++

	return sheetId, nil
}
