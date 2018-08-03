package schedule

import (
	"github.com/Luxurioust/excelize"

	"errors"
	"fmt"
	"path"
	"sort"
	"strconv"
	"strings"
	"time"

	"github.com/go-steven/sanle/utils"
)

// 从EXCEL排班表中读取员工排班记录
// 输入参数：
//     excel: EXCEL文件名称
//     sheetNames: 需要读取的EXCEL SHEET列表
// 返回：
//     map[string][]*UserSchedule: 员工排班记录, key: UserName
//     考勤开始日期/结束日期, format: YYYY-MM-DD
func ReadScheduleFromExcel(excel string, sheets []string) (*SchedulesData, error) {
	// 打开excel文件
	xlsx, err := excelize.OpenFile(excel)
	if err != nil {
		Logger.Error(err)
		return nil, err
	}

	ret := &SchedulesData{
		Data: make(map[string][]*Schedule),
	}
	dates := []string{}
	// 依次读取每一个SHEET
	for _, sheet := range sheets {
		// 读取SHEET排班区块元数据
		blocks, err := GetScheduleBlocksInfo(xlsx, sheet)
		if err != nil {
			Logger.Error(err)
			return nil, err
		}
		Logger.Debugf("blocks: %s", utils.Json(blocks))

		for _, block := range blocks {
			blockData, blockDates, err := ReadScheduleBlock(xlsx, sheet, block.StartRowId, block.EndRowId)
			if err != nil {
				Logger.Error(err)
				return nil, err
			}
			if len(blockData) == 0 || len(blockDates) == 0 {
				continue
			}

			for staffName, records := range blockData {
				v, ok := ret.Data[staffName]
				if !ok {
					v = []*Schedule{}
				}
				v = append(v, records...)
				sort.Slice(v, func(i, j int) bool {
					return v[i].StartTime < v[j].StartTime
				})
				ret.Data[staffName] = v
			}

			dates = append(dates, blockDates...)
		}
	}

	if len(dates) > 0 {
		sort.Strings(dates)
		ret.StartDate = dates[0]
		ret.EndDate = dates[len(dates)-1]
	}

	return ret, nil
}

// 排班区块元数据
type ScheduleBlock struct {
	StartRowId int // 区块起始行
	EndRowId   int // 区块结束行
}

const (
	SHEET_MAX_ROWS = 10000 // EXCEL 单个表格最大有效行数
	BLOCK_MAX_ROWS = 100   // 单个排班区块最大行数
)

// 读取排班区块元数据
func GetScheduleBlocksInfo(xlsx *excelize.File, sheet string) ([]*ScheduleBlock, error) {
	ret := []*ScheduleBlock{}

	var rowsInfo *ScheduleBlock
	for rowIdx := 1; rowIdx < SHEET_MAX_ROWS; rowIdx++ {
		val := xlsx.GetCellValue(sheet, fmt.Sprintf("A%d", rowIdx))
		if strings.Contains(val, WORK_SHOP_YANGZHA) || strings.Contains(val, WORK_SHOP_NANCE) || strings.Contains(val, WORK_SHOP_JIAHUI) {
			if rowsInfo != nil {
				rowsInfo.EndRowId = rowIdx - 1
				ret = append(ret, rowsInfo)
				rowsInfo = nil
			}

			rowsInfo = &ScheduleBlock{StartRowId: rowIdx}
		}
	}
	if rowsInfo != nil {
		rowsInfo.EndRowId = rowsInfo.StartRowId + BLOCK_MAX_ROWS
		ret = append(ret, rowsInfo)
	}

	return ret, nil
}

// 按照区块，读取排班数据
func ReadScheduleBlock(xlsx *excelize.File, sheet string, startRowId, endRowId int) (map[string][]*Schedule, []string, error) {
	// 读取区块中的工作店铺信息，必须在规定位置，否则报错
	workShop := xlsx.GetCellValue(sheet, fmt.Sprintf("A%d", startRowId))
	if workShop == "" {
		Logger.Errorf("No work shop, sheetName = %s", sheet)
		return nil, nil, nil
	}
	if strings.Contains(workShop, WORK_SHOP_YANGZHA) {
		workShop = WORK_SHOP_YANGZHA
	} else if strings.Contains(workShop, WORK_SHOP_NANCE) {
		workShop = WORK_SHOP_NANCE
	} else if strings.Contains(workShop, WORK_SHOP_JIAHUI) {
		workShop = WORK_SHOP_JIAHUI
	}
	Logger.Debugf("workShop = %s", utils.Json(workShop))

	// 读取区块中的排班年信息，必须在规定位置，否则报错
	scheduleYear := xlsx.GetCellValue(sheet, fmt.Sprintf("A%d", startRowId+1))
	if scheduleYear == "" {
		return nil, nil, errors.New("No schedule year")
	}
	Logger.Debugf("scheduleYear = %s", utils.Json(scheduleYear))

	columns := utils.GetExcelColumns(26 * 2)

	// 读取所有原始考勤日期
	oriDates := []string{}
	for idx, v := range columns {
		if idx == 0 {
			continue
		}
		val := xlsx.GetCellValue(sheet, fmt.Sprintf("%s%d", v, startRowId+1))
		if val != "" {
			oriDates = append(oriDates, val)
		} else {
			break
		}
	}
	Logger.Debugf("oriDates = %s", utils.Json(oriDates))

	loc, _ := time.LoadLocation("Asia/Chongqing")

	scheduleDateMap := make(map[string]struct{})
	data := map[string][]*Schedule{}
	for rowIdx := startRowId + 2; rowIdx <= endRowId; rowIdx++ {
		// 忽略前2行， 区块前两行是日期/星期标题
		staffName := xlsx.GetCellValue(sheet, fmt.Sprintf("A%d", rowIdx))
		if staffName == "" {
			continue // 忽略空行，排班区块中间允许有空行
		}

		for i, oriDate := range oriDates {
			vals := strings.Split(xlsx.GetCellValue(sheet, fmt.Sprintf("%s%d", columns[i+1], rowIdx)), "-")
			if len(vals) != 2 { // 排班时间格式必须是：{START}-{END}
				continue
			}

			// 格式化考勤日期格式：YYYY-MM-DD
			scheduleDate, err := FormatScheduleDate(scheduleYear, oriDate)
			if err != nil {
				Logger.Error(err)
				return nil, nil, err
			}

			// 格式化排班开始时间格式：YYYY-MM-DD HH24:MI
			startTime, startHour, err := FormatScheduleTime(vals[0])
			if err != nil {
				Logger.Error(err)
				return nil, nil, err
			}
			startTime = fmt.Sprintf("%s %s", scheduleDate, startTime)

			// 格式化排班结束时间格式：YYYY-MM-DD HH24:MI
			endTime, endHour, err := FormatScheduleTime(vals[1])
			if err != nil {
				Logger.Error(err)
				return nil, nil, err
			}
			if startHour > endHour { // 例如：22-7
				t, err := time.ParseInLocation("2006-01-02", scheduleDate, loc)
				if err != nil {
					Logger.Error(err)
					return nil, nil, err
				}
				endTime = fmt.Sprintf("%s %s", t.Add(24*time.Hour).Format("2006-01-02"), endTime)
			} else {
				if endHour == 24 { // 例如：18-24
					t, err := time.ParseInLocation("2006-01-02", scheduleDate, loc)
					if err != nil {
						Logger.Error(err)
						return nil, nil, err
					}
					endTime = fmt.Sprintf("%s 00:00", t.Add(24*time.Hour).Format("2006-01-02"))
				} else {
					endTime = fmt.Sprintf("%s %s", scheduleDate, endTime)
				}
			}

			// 判定是白班还是夜班
			rank := RANK_DAY
			if startHour > endHour || startHour < 5 {
				rank = RANK_NIGHT
			}

			record := &Schedule{
				StaffName:    staffName,
				ScheduleDate: scheduleDate,
				StartTime:    startTime,
				EndTime:      endTime,
				WorkShop:     workShop,
				Rank:         rank,
			}
			// 计算结算工时
			if err := CalcWageHours(record); err != nil {
				Logger.Error(err)
				return nil, nil, err
			}

			// 保存数据并排序
			records, ok := data[staffName]
			if !ok {
				records = []*Schedule{}
			}
			records = append(records, record)
			sort.Slice(records, func(i, j int) bool {
				return records[i].StartTime < records[j].StartTime
			})
			data[staffName] = records

			// 保存考勤日期
			if _, ok := scheduleDateMap[scheduleDate]; !ok {
				scheduleDateMap[scheduleDate] = struct{}{}
			}
		}
	}

	retDates := []string{}
	for k, _ := range scheduleDateMap {
		retDates = append(retDates, k)
	}
	sort.Strings(retDates)

	Logger.Debugf("data: %s", utils.Json(data))
	Logger.Debugf("scheduleDates: %s", utils.Json(retDates))
	return data, retDates, nil
}

// 格式化排班时间，输出格式：HH24:MI
// 输入参数取值示例：
//    8
//    8:30
//    12:00
//    23:15
// 输出：
//    HH24:MI格式的排班时间
//    对应的小时
func FormatScheduleTime(v string) (string, int, error) {
	vals := strings.Split(v, ":")
	if len(vals) == 1 {
		i, err := strconv.ParseInt(v, 10, 64)
		if err != nil {
			Logger.Error(err)
			return "", 0, err
		}
		return fmt.Sprintf("%02d:00", i), int(i), nil
	}

	i, err := strconv.ParseInt(vals[0], 10, 64)
	if err != nil {
		Logger.Error(err)
		return "", 0, err
	}
	return fmt.Sprintf("%02d:%s", i, vals[1]), int(i), nil
}

// 格式化排班日期，输出格式：YYYY-MM-DD
// 输入参数：
//      年
//      日期，取值示例：7/1, 6/30, 11/3
func FormatScheduleDate(year, scheduleDate string) (string, error) {
	loc, _ := time.LoadLocation("Asia/Chongqing")

	vals := strings.Split(scheduleDate, "/")
	if len(vals) != 2 {
		return "", fmt.Errorf("invalid schedule date: %s", scheduleDate)
	}
	month, err := strconv.ParseInt(vals[0], 10, 64)
	if err != nil {
		Logger.Error(err)
		return "", err
	}
	day, err := strconv.ParseInt(vals[1], 10, 64)
	if err != nil {
		Logger.Error(err)
		return "", err
	}

	t, err := time.ParseInLocation("2006/01/02", fmt.Sprintf("%s/%02d/%02d", year, month, day), loc)
	if err != nil {
		Logger.Error(err)
		return "", err
	}

	return t.Format("2006-01-02"), nil
}

const (
	WAGE_NIGHT_START_HOUR = 22
	WAGE_NIGHT_END_HOUR   = 7
)

// 排班记录：计算并更新工时结算
func CalcWageHours(record *Schedule) error {
	loc, _ := time.LoadLocation("Asia/Chongqing")

	// 考勤日期
	t_schedule_date, err := time.ParseInLocation("2006-01-02", record.ScheduleDate, loc)
	if err != nil {
		Logger.Error(err)
		return err
	}

	// 夜班工时对应开始时间
	t_day_last_time := t_schedule_date.Add(WAGE_NIGHT_START_HOUR * time.Hour)
	// 夜班工时对应结束时间
	t_night_last_time := t_schedule_date.Add((24 + WAGE_NIGHT_END_HOUR) * time.Hour)

	// 排班开始时间
	t_start, err := time.ParseInLocation("2006-01-02 15:04", record.StartTime, loc)
	if err != nil {
		Logger.Error(err)
		return err
	}

	// 排班结束时间
	t_end, err := time.ParseInLocation("2006-01-02 15:04", record.EndTime, loc)
	if err != nil {
		Logger.Error(err)
		return err
	}

	if record.Rank == RANK_DAY { // 白班
		var realDayHours, realNightHours float64
		if t_end.Before(t_day_last_time) { // 白班未过22点
			realDayHours = t_end.Sub(t_start).Minutes() / 60
		} else { // 白班超过22点
			realDayHours = t_day_last_time.Sub(t_start).Minutes() / 60
			realNightHours = t_end.Sub(t_day_last_time).Minutes() / 60
		}

		// 5小时以上（不包含5小时），扣除30分钟休息时间
		if realDayHours+realNightHours > 5.0 {
			if realDayHours > 0.5 {
				record.WageDayHours = realDayHours - 0.5
				record.WageNightHours = realNightHours
			} else {
				record.WageDayHours = realDayHours
				record.WageNightHours = realNightHours - 0.5
			}
		} else {
			record.WageDayHours = realDayHours
			record.WageNightHours = realNightHours
		}
	} else { // 夜班
		var (
			realDayHours_1, realDayHours_2       float64
			realNightStartTime, realNightEndTime time.Time
		)

		if t_start.Before(t_day_last_time) { // 夜班开始时间在22点之前
			realDayHours_1 = t_day_last_time.Sub(t_start).Minutes() / 60
			realNightStartTime = t_day_last_time
		} else {
			realNightStartTime = t_start
		}

		if t_end.After(t_night_last_time) { // 夜班结束时间在7点之后
			realDayHours_2 = t_end.Sub(t_night_last_time).Minutes() / 60
			realNightEndTime = t_night_last_time
		} else {
			realNightEndTime = t_end
		}
		realNightHours := realNightEndTime.Sub(realNightStartTime).Minutes() / 60
		realDayHours := realDayHours_1 + realDayHours_2

		// 5小时以上（不包含5小时），扣除30分钟休息时间
		if realDayHours+realNightHours > 5.0 {
			if realNightHours > 0.5 {
				record.WageDayHours = realDayHours
				record.WageNightHours = realNightHours - 0.5
			} else {
				record.WageDayHours = realDayHours - 0.5
				record.WageNightHours = realNightHours
			}
		} else {
			record.WageDayHours = realDayHours
			record.WageNightHours = realNightHours
		}
	}

	//Logger.Infof("record: %s", utils.Json(record))

	return nil
}

func GenerateOutputFilename(scheduleExcel string, subName string) string {
	vals := strings.Split(path.Base(scheduleExcel), ".")
	return fmt.Sprintf("%s/%s_%s%s", path.Dir(scheduleExcel), vals[0], subName, path.Ext(scheduleExcel))
}
