package main

import (
	"fmt"

	"github.com/miaomiao3/xlsx"
	"github.com/miaomiao3/xlsxutil"
)

const (
	xlsxFilePath  = "people.xlsx"
	xlsxFileSheet = "persons"
)

type Person struct {
	Name  string  `xls:"name"`
	Money float64 `xls:"money,precision:2"`
	Age   int     `xls:"age"`
	Edu   `xls:",inline"`
}

type Edu struct {
	School  string `xls:"school"`
	Address string `xls:"address"`
}

func main() {

	persons := make([]*Person, 0)
	file, err := xlsx.OpenFile(xlsxFilePath)
	if err != nil {
		panic(err)
	}

	err = xlsxutil.XlsLoad(file, xlsxFileSheet, &persons)
	fmt.Println("persons:", persons)

	file2 := xlsx.NewFile()
	err = xlsxutil.XlsDump(file2, xlsxFileSheet, persons)
	if err != nil {
		panic(err)
	}

	file2.Save("test.xlsx")
	// please view the file to check data

}
