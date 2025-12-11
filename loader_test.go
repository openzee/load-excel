package loader

import (
	"fmt"
	"testing"
)

func Print(allPoints []*Point) {

	for _, p := range allPoints {
		fmt.Printf("%s\n", p)
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

func TestD(t *testing.T) {

	rst, err := ParseExcel("test/test.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}

	rst2 := ReassembleWithGroupIDAndFreq(rst)
	for a, b := range rst2 {
		fmt.Println(a, b)
	}
}
