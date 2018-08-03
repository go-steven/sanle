package utils

import (
	"testing"
)

func TestGetMonthDays(t *testing.T) {
	for month, expected := range map[int]int{
		1: 31, 2: 28, 3: 31, 4: 30, 5: 31, 6: 30, 7: 31, 8: 31, 9: 30, 10: 31, 11: 30, 12: 31,
	} {
		ret := GetMonthDays(2018, month)
		if ret != expected {
			t.Errorf("GetMonthDays(%d, %d) was incorrect, got: %d, want: %d.", 2018, month, ret, expected)
			return
		}
	}

	for year, expected := range map[int]int{
		2018: 28, 2020: 29, 2200: 28, 2000: 29, // 四年一闰，百年不闰，四百年又闰
	} {
		ret := GetMonthDays(year, 2)
		if ret != expected {
			t.Errorf("GetMonthDays(%d, %d) was incorrect, got: %d, want: %d.", year, 2, ret, expected)
			return
		}
	}
}
