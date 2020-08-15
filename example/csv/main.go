package main

import (
	"fmt"
	"github.com/miaomiao3/xlsxutil"
)

const (
	csvFilePath = "people.csv"
)

type Person struct {
	Name  string  `xls:"name"`
	Money float64 `xls:"money,precision:2"`
	Age   int     `xls:"age"`
	Edu   `xls:",inline"`
}

type Edu struct {
	School  string `xls:"school",yaml:"school"`
	Address string `xls:"address",yaml:"address"`
}

func main() {

	persons := make([]*Person, 0)

	err := xlsxutil.CsvLoad(csvFilePath, ",", &persons)
	if err != nil {
		panic(err)
	}
	fmt.Println("persons:", persons)

	buf, err := xlsxutil.CsvDump(",", persons)
	if err != nil {
		panic(err)
	}

	fmt.Println("buf")
	fmt.Println(buf)
}
