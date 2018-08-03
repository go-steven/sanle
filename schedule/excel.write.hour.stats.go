package schedule

import (
	"github.com/Luxurioust/excelize"

	"fmt"
	"time"

	"github.com/go-steven/sanle/utils"
)

// 将工时分布统计数据写入EXCEL中
func SaveHourStatsToExcel(excel string, hourStatsData *HourStatData) error {
	xlsx := excelize.NewFile()

	workShops := []string{WORK_SHOP_YANGZHA, WORK_SHOP_NANCE, WORK_SHOP_JIAHUI}
	for _, v := range workShops {
		_, err := SaveShopHourStatsToExcel(xlsx, v, hourStatsData)
		if err != nil {
			Logger.Error(err)
			return err
		}
	}
	// Set active sheet of the workbook.
	xlsx.SetActiveSheet(2)
	// Save xlsx file by the given path.
	if err := xlsx.SaveAs(excel); err != nil {
		Logger.Error(err)
		return err
	}

	return nil
}

const HOUR_STATS_MONTH_ROWS int = 51 + 1

func SaveShopHourStatsToExcel(xlsx *excelize.File, workShop string, hourStatsData *HourStatData) (int, error) {
	sheetName := workShop
	sheetId := xlsx.NewSheet(sheetName)

	loc, _ := time.LoadLocation("Asia/Chongqing")

	//Logger.Infof("startDate : %s, endDate: %s", startDate, endDate)
	t_start, err := time.ParseInLocation("2006-01-02", hourStatsData.StartDate, loc)
	if err != nil {
		Logger.Error(err)
		return 0, err
	}

	t_end, err := time.ParseInLocation("2006-01-02", hourStatsData.EndDate, loc)
	if err != nil {
		Logger.Error(err)
		return 0, err
	}

	if t_end.Before(t_start.Add(30*24*time.Hour)) || (t_start.Year() == t_end.Year() && t_start.Month() == t_end.Month()) { // 30天之内 或者 同一个月
		if err := write_hour_stats_block(xlsx, sheetName, 0, workShop, hourStatsData.Data[workShop], hourStatsData.StartDate, hourStatsData.EndDate); err != nil {
			return 0, err
		}
	} else {
		dates := []time.Time{}
		blockCnt := 0
		t_curr_date := t_start
		for {
			dates = append(dates, t_curr_date)

			t_curr_date = t_curr_date.Add(24 * time.Hour)
			if t_curr_date.After(t_end) {
				break
			}

			if t_curr_date.Month() != t_start.Month() || t_curr_date.Year() != t_start.Year() { // 按自然月写入工时分布数据
				if len(dates) > 0 {
					if err := write_hour_stats_block(xlsx, sheetName, blockCnt*HOUR_STATS_MONTH_ROWS, workShop, hourStatsData.Data[workShop], dates[0].Format("2006-01-02"), dates[len(dates)-1].Format("2006-01-02")); err != nil {
						return 0, err
					}
					dates = []time.Time{}
					blockCnt++
				}
				t_start = t_curr_date
			}
		}
		if len(dates) > 0 {
			if err := write_hour_stats_block(xlsx, sheetName, blockCnt*HOUR_STATS_MONTH_ROWS, workShop, hourStatsData.Data[workShop], dates[0].Format("2006-01-02"), dates[len(dates)-1].Format("2006-01-02")); err != nil {
				return 0, err
			}
		}
	}

	return sheetId, nil
}

func write_hour_stats_block(xlsx *excelize.File, sheet string, startRowId int, workShop string, hourStats map[string]int, startDate, endDate string) error {
	loc, _ := time.LoadLocation("Asia/Chongqing")

	t_start, err := time.ParseInLocation("2006-01-02", startDate, loc)
	if err != nil {
		Logger.Error(err)
		return err
	}

	t_end, err := time.ParseInLocation("2006-01-02", endDate, loc)
	if err != nil {
		Logger.Error(err)
		return err
	}

	dates := []time.Time{}
	t_date := t_start
	for {
		dates = append(dates, t_date)

		t_date = t_date.Add(24 * time.Hour)
		if t_date.After(t_end) {
			break
		}
	}

	columns := utils.GetExcelColumns(26 * 2)

	xlsx.SetCellValue(sheet, fmt.Sprintf("A%d", startRowId+2), workShop)
	xlsx.SetCellValue(sheet, fmt.Sprintf("A%d", startRowId+3), "工时小计")

	for idx, v := range dates {
		xlsx.SetCellValue(sheet, fmt.Sprintf("%s%d", columns[idx+1], startRowId+1), utils.GetWeekDay(v))
		xlsx.SetCellValue(sheet, fmt.Sprintf("%s%d", columns[idx+1], startRowId+2), fmt.Sprintf("%d/%d", v.Month(), v.Day()))
	}

	for i := 0; i < 48; i++ {
		hour := i / 2
		var startHour, endHour string
		if i%2 == 1 {
			startHour = fmt.Sprintf("%02d:30", hour)
			endHour = fmt.Sprintf("%02d:00", hour+1)
		} else {
			startHour = fmt.Sprintf("%02d:00", hour)
			endHour = fmt.Sprintf("%02d:30", hour)
		}
		key := fmt.Sprintf("%s_%s", startHour, endHour)
		xlsx.SetCellValue(sheet, fmt.Sprintf("A%d", startRowId+i+2+2), key)
	}

	for idx, v := range dates {
		dayHours := 0
		for i := 0; i < 48; i++ {
			hour := i / 2
			var startHour, endHour string
			if i%2 == 1 {
				startHour = fmt.Sprintf("%s %02d:30", v.Format("2006-01-02"), hour)
				endHour = fmt.Sprintf("%s %02d:00", v.Format("2006-01-02"), hour+1)
			} else {
				startHour = fmt.Sprintf("%s %02d:00", v.Format("2006-01-02"), hour)
				endHour = fmt.Sprintf("%s %02d:30", v.Format("2006-01-02"), hour)
			}
			key := fmt.Sprintf("%s_%s", startHour, endHour)
			val, ok := hourStats[key]
			if !ok {
				val = 0
			}
			dayHours += val
			cell := fmt.Sprintf("%s%d", columns[idx+1], startRowId+i+2+2)
			xlsx.SetCellValue(sheet, cell, val)
		}
		xlsx.SetCellValue(sheet, fmt.Sprintf("%s%d", columns[idx+1], startRowId+3), dayHours/2)
	}

	return nil
}
