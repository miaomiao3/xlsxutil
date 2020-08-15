## xlsxutil

Helpful func to read/write csv or xlsx files.

# Feature
* Easy api. Only `*Load` and `*Dump` function.
* Inline support

# Limit
* Only support string and numeric data type

***
## Usage
Here a new tag 'xls' is used for this repo.  
Pay attention to keys:
* `inline` for extract struct
* `precision` for `floats` data type precision control.

Refer to example for mor cases.
 


# Quickstart

Get codes  
` $ go get github.com/miaomiao3/xlsxutil`

Import  
`import ( "github.com/miaomiao3/xlsxutil" )`

***
supposed a csv document content:

```$xslt
name,money,age,school,address
n-0,1.2345678900,20,school-0,hali-0
n-1,1.2345678900,21,school-1,hali-1
n-2,1.2345678900,22,school-2,hali-2
n-3,1.2345678900,23,school-3,hali-3
n-4,1.2345678900,24,school-4,hali-4
```
To bind this csv to out data, and dump data to csv string, just like

```go
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

```
terminal output:
```$xslt
$ go run main.go
persons: [0xc000078140 0xc000078180 0xc0000781c0 0xc000078200 0xc000078240]
buf
name,money,age,school,address
n-1,1.23,21,school-1,hali-1
n-2,1.23,22,school-2,hali-2
n-3,1.23,23,school-3,hali-3
n-4,1.23,24,school-4,hali-4

```

# Examples
For other examples, Please refer to [examples](https://github.com/miaomiao3/xlsxutil/tree/master/example) dir.



