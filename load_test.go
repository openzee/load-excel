package excel

import (
	"fmt"
	"testing"
)

func Print(allPoints []*Point) {

	for _, p := range allPoints {
		fmt.Println(*p)
	}
}

func TestA(t *testing.T) {
	rst, err := ParseExcel("test/test.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}

	Print(rst)
}

func TestB(t *testing.T) {
	rst, err := ParseExcel("test/test.xlsx", []string{"Sheet2", "Sheet1"}...)
	if err != nil {
		fmt.Println(err)
		return
	}

	Print(rst)
}

func TestC(t *testing.T) {

	rst, err := ParseExcel("test/test.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}

	rst2 := ReassembleWithAddrAndFreq(rst)
	fmt.Println(rst2)
}
