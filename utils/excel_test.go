package utils

import (
	"testing"
)

func TestGetExcelColumns(t *testing.T) {
	data := [][]string{
		[]string{"A", "B", "C"},
		[]string{"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"},
		[]string{"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD"},
		[]string{"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "BA"},
	}
	for _, expected := range data {
		ret := GetExcelColumns(len(expected))
		if len(ret) != len(expected) {
			t.Errorf("GetExcelColumns(%d) was incorrect, got len(ret): %d, want: %d.", len(expected), len(ret), len(expected))
			return
		}
		if len(ret) > 0 {
			for i, v := range ret {
				if v != expected[i] {
					t.Errorf("GetExcelColumns(%d) was incorrect, got ret[%d]: %s, want: %s.", len(expected), i, v, expected[i])
					return
				}
			}
		}
	}
}
