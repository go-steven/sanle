package kaoqing

import (
	"fmt"
	"math"
	"sort"
	"time"

	"github.com/go-steven/sanle/schedule"
)

func UpdateUserKaoQing(schedules map[string][]*schedule.Schedule, origRecords map[string][]*OrigRecord, startDate, endDate string) (map[string][]*KaoQing, error) {
	origMap := make(map[string][]*OrigRecord)
	for _, rows := range origRecords {
		for _, v := range rows {
			key := fmt.Sprintf("%s_%s_%s", v.StaffName, v.WorkShop, v.RecordDate)
			val, ok := origMap[key]
			if !ok {
				val = []*OrigRecord{}
			}
			val = append(val, v)
			sort.Slice(val, func(i, j int) bool {
				return val[i].RecordOn < val[j].RecordOn
			})
			origMap[key] = val
		}
	}

	loc, _ := time.LoadLocation("Asia/Chongqing")
	usedOrigMap := make(map[string]struct{})
	today := time.Now().Format("2006-01-02")

	ret := make(map[string][]*KaoQing)
	for staffName, rows := range schedules {
		kaoqingList := []*KaoQing{}
		for _, s := range rows {
			if startDate != "" && s.ScheduleDate < startDate {
				continue
			}
			if endDate != "" && s.ScheduleDate > endDate {
				continue
			}

			if s.ScheduleDate >= today {
				continue
			}

			kaoqing := &KaoQing{
				Schedule: schedule.Schedule{
					StaffName:    s.StaffName,
					ScheduleDate: s.ScheduleDate,
					StartTime:    s.StartTime,
					EndTime:      s.EndTime,
					Rank:         s.Rank,
					WorkShop:     s.WorkShop,
				},
			}

			t_date, err := time.ParseInLocation("2006-01-02", s.ScheduleDate, loc)
			if err != nil {
				Logger.Error(err)
				return nil, err
			}

			t_start, err := time.ParseInLocation("2006-01-02 15:04", s.StartTime, loc)
			if err != nil {
				Logger.Error(err)
				return nil, err
			}

			t_end, err := time.ParseInLocation("2006-01-02 15:04", s.EndTime, loc)
			if err != nil {
				Logger.Error(err)
				return nil, err
			}

			if kaoqing.StartRecord == nil {
				subOrigRecords := []*OrigRecord{}
				if subOrigRecordsA, ok := origMap[fmt.Sprintf("%s_%s_%s", s.StaffName, s.WorkShop, s.ScheduleDate)]; ok {
					subOrigRecords = append(subOrigRecords, subOrigRecordsA...)
				}
				if subOrigRecordsB, ok := origMap[fmt.Sprintf("%s_%s_%s", s.StaffName, s.WorkShop, t_date.Add((-1)*24*time.Hour).Format("2006-01-02"))]; ok {
					subOrigRecords = append(subOrigRecords, subOrigRecordsB...)
				}

				if len(subOrigRecords) > 0 {
					sort.Slice(subOrigRecords, func(i, j int) bool {
						return subOrigRecords[i].RecordOn < subOrigRecords[j].RecordOn
					})

					for _, v := range subOrigRecords {
						if _, ok := usedOrigMap[fmt.Sprintf("%s_%s_%s", v.StaffName, v.WorkShop, v.RecordOn)]; ok {
							continue
						}

						t_record_on, err := time.ParseInLocation("2006-01-02 15:04", v.RecordOn, loc)
						if err != nil {
							Logger.Error(err)
							return nil, err
						}

						if t_record_on.Equal(t_start) || (t_record_on.Before(t_start) && t_start.Sub(t_record_on) < 2*time.Hour) || (t_start.Before(t_record_on) && t_record_on.Sub(t_start) < 2*time.Hour) {
							kaoqing.StartRecord = v
							usedOrigMap[fmt.Sprintf("%s_%s_%s", v.StaffName, v.WorkShop, v.RecordOn)] = struct{}{}
							break
						}
					}
				}
			}

			if kaoqing.EndRecord == nil {
				subOrigRecords := []*OrigRecord{}
				if subOrigRecordsA, ok := origMap[fmt.Sprintf("%s_%s_%s", s.StaffName, s.WorkShop, s.ScheduleDate)]; ok {
					subOrigRecords = append(subOrigRecords, subOrigRecordsA...)
				}
				if subOrigRecordsB, ok := origMap[fmt.Sprintf("%s_%s_%s", s.StaffName, s.WorkShop, t_date.Add(24*time.Hour).Format("2006-01-02"))]; ok {
					subOrigRecords = append(subOrigRecords, subOrigRecordsB...)
				}
				if len(subOrigRecords) > 0 {
					sort.Slice(subOrigRecords, func(i, j int) bool {
						return subOrigRecords[i].RecordOn < subOrigRecords[j].RecordOn
					})

					for _, v := range subOrigRecords {
						if _, ok := usedOrigMap[fmt.Sprintf("%s_%s_%s", v.StaffName, v.WorkShop, v.RecordOn)]; ok {
							continue
						}

						t_record_on, err := time.ParseInLocation("2006-01-02 15:04", v.RecordOn, loc)
						if err != nil {
							Logger.Error(err)
							return nil, err
						}

						if t_record_on.Equal(t_end) || (t_record_on.Before(t_end) && t_end.Sub(t_record_on) < 3*time.Hour) || (t_end.Before(t_record_on) && t_record_on.Sub(t_end) < 3*time.Hour) {
							kaoqing.EndRecord = v
							usedOrigMap[fmt.Sprintf("%s_%s_%s", v.StaffName, v.WorkShop, v.RecordOn)] = struct{}{}
							break
						}
					}
				}
			}

			if kaoqing.StartRecord == nil {
				if kaoqing.EndRecord == nil {
					kaoqing.RecordSts = RECORD_STS_NONE
				} else {
					kaoqing.RecordSts = RECORD_STS_NOSTART
				}
			} else if kaoqing.EndRecord == nil {
				kaoqing.RecordSts = RECORD_STS_NOEND
			} else {
				kaoqing.RecordSts = RECORD_STS_NORMAL
			}

			var lateMinutes, overMinutes int
			if kaoqing.StartRecord != nil {
				lateMinutes, err = get_minutes(kaoqing.StartRecord.RecordOn, s.StartTime)
				if err != nil {
					Logger.Error(err)
					return nil, err
				}
			}
			if kaoqing.EndRecord != nil {
				overMinutes, err = get_minutes(kaoqing.EndRecord.RecordOn, s.EndTime)
				if err != nil {
					Logger.Error(err)
					return nil, err
				}
			}
			if overMinutes < -3 && lateMinutes > 3 { // 迟到超过3分钟算迟到
				kaoqing.Sts = KAOQING_STS_LATE_AND_EARLY
				kaoqing.LateMinutes = lateMinutes
				kaoqing.EarlyMinutes = (-1) * overMinutes
			} else if overMinutes < -3 { // 早退3分钟也算早退
				kaoqing.Sts = KAOQING_STS_EARLY
				kaoqing.EarlyMinutes = (-1) * overMinutes
			} else if lateMinutes > 3 {
				kaoqing.Sts = KAOQING_STS_LATE
				kaoqing.LateMinutes = lateMinutes
			} else {
				if kaoqing.RecordSts == RECORD_STS_NORMAL {
					kaoqing.Sts = KAOQING_STS_NORMAL
				}
			}

			if overMinutes > 30 { // 加班超过30分钟算加班
				kaoqing.OverMinutes = overMinutes
			}

			kaoqingList = append(kaoqingList, kaoqing)
		}

		// 保存数据并排序
		data, ok := ret[staffName]
		if !ok {
			data = []*KaoQing{}
		}
		data = append(data, kaoqingList...)
		sort.Slice(data, func(i, j int) bool {
			return data[i].StartTime < data[j].StartTime
		})
		ret[staffName] = data
	}

	return ret, nil
}

// 取分钟差
func get_minutes(aTime, bTime string) (int, error) {
	loc, _ := time.LoadLocation("Asia/Chongqing")

	t_a, err := time.ParseInLocation("2006-01-02 15:04", aTime, loc)
	if err != nil {
		Logger.Error(err)
		return 0, err
	}

	t_b, err := time.ParseInLocation("2006-01-02 15:04", bTime, loc)
	if err != nil {
		Logger.Error(err)
		return 0, err
	}

	return int(math.Floor(t_a.Sub(t_b).Minutes())), nil
}
