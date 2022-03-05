package main

import (
	"encoding/json"
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
	School  string `xls:"school"`
	Address string `xls:"address"`
}

func main() {

	persons := make([]*Person, 0)

	err := xlsxutil.CsvLoad(csvFilePath, ",", &persons)
	if err != nil {
		fmt.Printf("%+v", err)
		panic(err)
	}
	fmt.Println("persons:", GetJsonIndent(persons))

	//buf, err := xlsxutil.CsvDump(",", persons)
	//if err != nil {
	//	panic(err)
	//}
	//
	//fmt.Println("buf")
	//fmt.Println(buf)
}
func GetJsonIndent(v interface{}) string {
	out, err := json.MarshalIndent(v, "", "  ")
	if err != nil {
		return err.Error()
	} else {
		return string(out)
	}
}
