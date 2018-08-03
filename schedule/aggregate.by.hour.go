package schedule

import (
	"fmt"
	"time"
)

// 根据员工排班记录，生成每天每小时的工时分布统计
func AggregateByHour(schedulesData *SchedulesData) (*HourStatData, error) {
	ret := &HourStatData{
		Data:      make(map[string]map[string]int),
		StartDate: schedulesData.StartDate,
		EndDate:   schedulesData.EndDate,
	}

	for _, records := range schedulesData.Data {
		for _, record := range records {
			if _, ok := ret.Data[record.WorkShop]; !ok {
				ret.Data[record.WorkShop] = make(map[string]int)
			}

			hourStats, err := GetHourStatsByRecord(record)
			if err != nil {
				Logger.Error(err)
				return nil, err
			}
			for k, v := range hourStats {
				if _, ok := ret.Data[record.WorkShop][k]; !ok {
					ret.Data[record.WorkShop][k] = v
				} else {
					ret.Data[record.WorkShop][k] += v
				}
			}

		}
	}

	return ret, nil
}

// 单条排班记录中的工时分布统计
func GetHourStatsByRecord(record *Schedule) (map[string]int, error) {
	loc, _ := time.LoadLocation("Asia/Chongqing")

	t_start, err := time.ParseInLocation("2006-01-02 15:04", record.StartTime, loc)
	if err != nil {
		Logger.Error(err)
		return nil, err
	}

	t_end, err := time.ParseInLocation("2006-01-02 15:04", record.EndTime, loc)
	if err != nil {
		Logger.Error(err)
		return nil, err
	}

	ret := make(map[string]int)
	if t_end.Format("2006-01-02") == record.ScheduleDate {
		for i := t_start.Hour(); i < t_end.Hour(); i++ {
			if i == t_start.Hour() && t_start.Minute() > 0 {
				startHour := record.StartTime
				endHour := fmt.Sprintf("%s %02d:00", record.ScheduleDate, i+1)
				ret[fmt.Sprintf("%s_%s", startHour, endHour)] = 1
			} else {
				startHour := fmt.Sprintf("%s %02d:00", record.ScheduleDate, i)
				endHour := fmt.Sprintf("%s %02d:30", record.ScheduleDate, i)
				ret[fmt.Sprintf("%s_%s", startHour, endHour)] = 1

				startHour = fmt.Sprintf("%s %02d:30", record.ScheduleDate, i)
				endHour = fmt.Sprintf("%s %02d:00", record.ScheduleDate, i+1)
				ret[fmt.Sprintf("%s_%s", startHour, endHour)] = 1
			}
		}
	} else {
		for i := t_start.Hour(); i < 24; i++ {
			if i == t_start.Hour() && t_start.Minute() > 0 {
				startHour := record.StartTime
				endHour := fmt.Sprintf("%s %02d:00", record.ScheduleDate, i+1)
				ret[fmt.Sprintf("%s_%s", startHour, endHour)] = 1
			} else {
				startHour := fmt.Sprintf("%s %02d:00", record.ScheduleDate, i)
				endHour := fmt.Sprintf("%s %02d:30", record.ScheduleDate, i)
				ret[fmt.Sprintf("%s_%s", startHour, endHour)] = 1

				startHour = fmt.Sprintf("%s %02d:30", record.ScheduleDate, i)
				endHour = fmt.Sprintf("%s %02d:00", record.ScheduleDate, i+1)
				ret[fmt.Sprintf("%s_%s", startHour, endHour)] = 1
			}
		}
		next_date := t_end.Format("2006-01-02")
		for i := 0; i < t_end.Hour(); i++ {
			startHour := fmt.Sprintf("%s %02d:00", next_date, i)
			endHour := fmt.Sprintf("%s %02d:30", next_date, i)
			ret[fmt.Sprintf("%s_%s", startHour, endHour)] = 1

			startHour = fmt.Sprintf("%s %02d:30", next_date, i)
			endHour = fmt.Sprintf("%s %02d:00", next_date, i+1)
			ret[fmt.Sprintf("%s_%s", startHour, endHour)] = 1
		}
	}
	if t_end.Minute() > 0 {
		startHour := t_end.Format("2006-01-02 15") + ":00"
		endHour := record.EndTime
		ret[fmt.Sprintf("%s_%s", startHour, endHour)] = 1
	}

	return ret, nil
}
