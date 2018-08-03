package utils

import (
	"fmt"
)

func GetExcelColumns(columnCnt int) []string {
	baseList := []string{"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}
	baseCnt := len(baseList)
	if columnCnt <= baseCnt {
		return baseList[0:columnCnt]
	}

	var loop int = columnCnt / baseCnt
	var left int = columnCnt % baseCnt
	var prefix string
	ret := []string{}
	for i := 0; i <= loop; i++ {
		if i > 0 {
			prefix = baseList[i-1]
		}
		for k, v := range baseList {
			if i < loop || k < left {
				ret = append(ret, fmt.Sprintf("%s%s", prefix, v))
			}
		}
	}

	return ret
}
