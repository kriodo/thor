package importer

import (
	"math"
	"strconv"
	"testing"
)

func TestEstimateRowCount(t *testing.T) {
	rate := float64(111) / float64(222) * 100
	processRate := strconv.FormatFloat(math.Floor(rate*100)/100, 'f', 0, 64) + "%"
	t.Log(processRate)
	//file := "../../../tmp/big.xlsx"
	//sheet := "Sheet1"

	//start1 := time.Now()
	//count1, err := FastRowCount1(file, sheet)
	//if err != nil {
	//	t.Log("1.估算失败:", err)
	//	return
	//}
	//duration1 := time.Since(start1)
	//t.Logf("1.Excel大致数据行数: %d 耗时:%+v", count1, duration1)
	//start2 := time.Now()
	//count2, err := FastRowCount2(file, sheet)
	//if err != nil {
	//	t.Log("2.估算失败:", err)
	//	return
	//}
	//duration2 := time.Since(start2)
	//t.Logf("2.Excel大致数据行数: %d 耗时:%+v", count2, duration2)

	//start3 := time.Now()
	//count3, err := FastRowCount3(file, sheet)
	//if err != nil {
	//	t.Log("3.估算失败:", err)
	//	return
	//}
	//duration3 := time.Since(start3)
	//t.Logf("3.Excel大致数据行数: %d 耗时:%+v", count3, duration3)
	//
	//start4 := time.Now()
	//count4, err := FastRowCount4(file, sheet)
	//if err != nil {
	//	t.Log("4.估算失败:", err)
	//	return
	//}
	//duration4 := time.Since(start4)
	//t.Logf("4.Excel大致数据行数: %d 耗时:%+v", count4, duration4)

	//start5 := time.Now()
	//count5, err := FastRowCount5(file, sheet)
	//if err != nil {
	//	t.Log("5.估算失败:", err)
	//	return
	//}
	//duration5 := time.Since(start5)
	//t.Logf("5.Excel大致数据行数: %d 耗时:%+v", count5, duration5)
	//t.Log("success")
}
