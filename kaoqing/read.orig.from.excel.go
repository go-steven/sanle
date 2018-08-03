package kaoqing

import (
	"github.com/Luxurioust/excelize"

	"fmt"
	"sort"
	"strings"
	"time"

	"github.com/go-steven/sanle/schedule"
)

const (
	ORIG_RECORD_MAX_ROWS    = 10000    // 最大读取记录数
	ORIG_RECORD_SHEET_NAME  = "Sheet3" // 原始记录对应的SHEET名称
	ORIG_RECORD_IGNORE_ROWS = 3        // 原始记录文件标题行数

	ORIG_RECORD_COLUMN_STAFF_NAME = "A" // EXCEL列名：员工姓名
	ORIG_RECORD_COLUMN_RECORD_ON  = "H" // EXCEL列名：打卡时间
	ORIG_RECORD_COLUMN_WORK_SHOP  = "J" // EXCEL列名：工作店铺
)

// 从EXCEL中读取原始打卡记录
func ReadOrignRecordFromExcel(excel string, startDate, endDate string) (map[string][]*OrigRecord, error) {
	loc, _ := time.LoadLocation("Asia/Chongqing")

	t_start, err := time.ParseInLocation("2006-01-02", startDate, loc)
	if err != nil {
		Logger.Error(err)
		return nil, err
	}
	startDate = t_start.Add((-1) * 24 * time.Hour).Format("2006-01-02")

	t_end, err := time.ParseInLocation("2006-01-02", endDate, loc)
	if err != nil {
		Logger.Error(err)
		return nil, err
	}
	endDate = t_end.Add(24 * time.Hour).Format("2006-01-02")

	// 打开EXCEL文件
	xlsx, err := excelize.OpenFile(excel)
	if err != nil {
		Logger.Error(err)
		return nil, err
	}
	sheet := ORIG_RECORD_SHEET_NAME

	ret := make(map[string][]*OrigRecord)
	for rowIdx := ORIG_RECORD_IGNORE_ROWS + 1; rowIdx < ORIG_RECORD_MAX_ROWS; rowIdx++ {
		// 忽略前3行
		staffName := xlsx.GetCellValue(sheet, fmt.Sprintf("%s%d", ORIG_RECORD_COLUMN_STAFF_NAME, rowIdx)) // 员工姓名
		if staffName != "" {                                                                              // 忽略空行
			recordOn := xlsx.GetCellValue(sheet, fmt.Sprintf("%s%d", ORIG_RECORD_COLUMN_RECORD_ON, rowIdx)) // 打卡时间
			recordDate := recordOn[0:10]
			if startDate != "" && recordDate < startDate {
				continue
			}
			if endDate != "" && recordDate > endDate {
				continue
			}

			workShop := xlsx.GetCellValue(sheet, fmt.Sprintf("%s%d", ORIG_RECORD_COLUMN_WORK_SHOP, rowIdx)) // 工作店铺

			// 工作店铺名称标准化
			if strings.Contains(workShop, schedule.WORK_SHOP_YANGZHA) {
				workShop = schedule.WORK_SHOP_YANGZHA
			} else if strings.Contains(workShop, schedule.WORK_SHOP_NANCE) {
				workShop = schedule.WORK_SHOP_NANCE
			} else if strings.Contains(workShop, schedule.WORK_SHOP_JIAHUI) {
				workShop = schedule.WORK_SHOP_JIAHUI
			}

			// 生成原始记录
			record := &OrigRecord{
				StaffName:  staffName,
				RecordDate: recordDate,
				RecordOn:   recordOn,
				WorkShop:   workShop,
			}

			// 保存原始记录并排序
			v, ok := ret[staffName]
			if !ok {
				ret[staffName] = []*OrigRecord{record}
			} else {
				v = append(v, record)
				sort.Slice(v, func(i, j int) bool {
					return v[i].RecordOn < v[j].RecordOn
				})
				ret[staffName] = v
			}
		}
	}

	return ret, nil
}
